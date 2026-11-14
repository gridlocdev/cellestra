import { DateConversion } from '../../formulas/evaluation/dateConversion.ts';
import { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { CellStyles, NumberFormats } from '../../styling/cellStyle.ts';
import { CellValue } from '../../values/cellValue.ts';
import { CsvOptions } from './csvOptions.ts';

// A double holds 15 significant digits exactly; a longer digit string is an ID, not a number.
const MAX_SIGNIFICANT_DIGITS = 15;
// LibreOffice Calc's two-digit-year window (PRD §8.3): 00–29 are 2000–2029, 30–99 are 1930–1999.
const TWO_DIGIT_YEAR_PIVOT = 29;

/**
 * Parse CSV text into rows of raw string fields per PRD §8.4: a leading BOM is stripped, CR, LF, and
 * CRLF all end a record, and RFC 4180 quoting applies (doubled quotes inside quoted fields).
 */
export function parseCsv(text: string, options: CsvOptions = CsvOptions.default): string[][] {
  const delimiter = options.delimiter;
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);

  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let fieldStarted = false;

  const endField = (): void => {
    row.push(field);
    field = '';
    fieldStarted = false;
  };
  const endRow = (): void => {
    endField();
    rows.push(row);
    row = [];
  };

  for (let i = 0; i < text.length; i++) {
    const c = text[i]!;
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += c;
      }
      continue;
    }

    if (c === '"' && (!fieldStarted || field.length === 0)) {
      inQuotes = true;
      fieldStarted = true;
    } else if (c === '\r') {
      if (text[i + 1] === '\n') i++;
      endRow();
    } else if (c === '\n') {
      endRow();
    } else if (c === delimiter) {
      endField();
    } else {
      field += c;
      fieldStarted = true;
    }
  }

  // Flush the final field/row unless the text ended exactly on a newline.
  if (field.length > 0 || row.length > 0 || fieldStarted) endRow();
  return rows;
}

/** Parse CSV and infer typed cell values into a new single-sheet workbook. */
export function readCsv(text: string, options: CsvOptions = CsvOptions.default): Workbook {
  const workbook = Workbook.createNew();
  fillFromCsv(workbook, workbook.activeSheet, text, options);
  return workbook;
}

/**
 * Fill an existing worksheet from CSV, inferring numbers and dates with the locale (the options',
 * else the workbook's effective locale). Everything else stays text, formula-like text included (as
 * in LibreOffice's default import); empty fields stay blank.
 */
export function fillFromCsv(
  workbook: Workbook,
  sheet: Worksheet,
  text: string,
  options: CsvOptions = CsvOptions.default,
): void {
  const locale = options.locale ?? workbook.settings.effectiveLocale;
  const rows = parseCsv(text, options);
  const numbers = new LocaleNumberParser(locale);
  const dates = new LocaleDateParser(locale);
  let dateStyleId: number | null = null;

  for (let r = 0; r < rows.length; r++) {
    const cells = rows[r]!;
    for (let c = 0; c < cells.length; c++) {
      const raw = cells[c]!;
      if (raw.length === 0) continue;

      const number = numbers.parse(raw);
      if (number !== null && Number.isFinite(number)) {
        sheet.setValue(r, c, CellValue.number(number));
        continue;
      }

      const date = dates.parse(raw);
      if (date !== null) {
        dateStyleId ??= workbook.styles.intern({
          ...CellStyles.default,
          numberFormat: NumberFormats.of('yyyy-mm-dd'),
        });
        sheet.setValue(r, c, CellValue.number(date));
        sheet.getOrCreate(r, c).styleId = dateStyleId;
        continue;
      }

      sheet.setValue(r, c, CellValue.text(raw));
    }
  }
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * `double.TryParse(raw, NumberStyles.Any, culture)`: surrounding whitespace, a leading or trailing
 * sign, parentheses for a negative, the locale's group and decimal separators, and an exponent.
 * Identifiers that merely look numeric are refused so they stay text: a leading zero ("00123") and
 * more than 15 significant digits (a 19-digit order ID would lose its last digits as a double).
 */
class LocaleNumberParser {
  private readonly pattern: RegExp;
  private readonly group: string;
  private readonly decimal: string;

  constructor(locale: string) {
    let group = ',';
    let decimal = '.';
    // Digits per group: the last group (primary) and the ones before it (3 and 3, or en-IN's 3 and 2).
    let primary = 3;
    let secondary = 3;
    try {
      const integers: string[] = [];
      for (const part of new Intl.NumberFormat(locale).formatToParts(1234567.5)) {
        if (part.type === 'group') group = part.value;
        else if (part.type === 'decimal') decimal = part.value;
        else if (part.type === 'integer') integers.push(part.value);
      }
      if (integers.length >= 2) primary = integers.at(-1)!.length;
      if (integers.length >= 3) secondary = integers.at(-2)!.length;
    } catch {
      // An unknown locale falls back to the invariant separators.
    }
    // Narrow/non-breaking spaces used as group separators also match a plain space.
    if (/\s/.test(group)) group = ' ';
    this.group = group;
    this.decimal = decimal;
    const g = escapeRegex(group);
    const d = escapeRegex(decimal);
    // A group separator only where groups fall: "1,234" and "12,345,678", never "1,2" (two values
    // to a reader in a comma locale, not twelve).
    const grouped = `\\d{1,${secondary}}(?:${g}\\d{${secondary}})*${g}\\d{${primary}}`;
    this.pattern = new RegExp(`^(${grouped}|\\d+)?(?:${d}(\\d*))?(?:[eE]([+-]?\\d+))?$`);
  }

  parse(raw: string): number | null {
    let t = raw.trim();
    if (this.group === ' ') t = t.replace(/[\u00a0\u202f]/g, ' ');
    let negative = false;

    if (t.startsWith('(') && t.endsWith(')')) {
      negative = true;
      t = t.slice(1, -1).trim();
    }
    if (t.startsWith('-') || t.startsWith('+')) {
      negative = negative !== (t[0] === '-');
      t = t.slice(1).trim();
    } else if (t.endsWith('-') || t.endsWith('+')) {
      negative = negative !== t.endsWith('-');
      t = t.slice(0, -1).trim();
    }
    if (t.length === 0) return null;

    const m = this.pattern.exec(t);
    if (!m) return null;
    const integer = (m[1] ?? '').split(this.group).join('');
    const fraction = m[2] ?? '';
    if (integer.length === 0 && fraction.length === 0) return null;
    if (/^0\d/.test(integer)) return null;
    if (significantDigits(integer + fraction) > MAX_SIGNIFICANT_DIGITS) return null;
    const exponent = m[3] === undefined ? '' : 'e' + m[3];
    const n = Number(`${integer || '0'}.${fraction || '0'}${exponent}`);
    return negative ? -n : n;
  }
}

function significantDigits(digits: string): number {
  return digits.replace(/^0+/, '').replace(/0+$/, '').length;
}

type DateField = 'day' | 'month' | 'year';

/**
 * `DateTime.TryParse(raw, culture)` for the shapes CSV exports actually carry: ISO
 * `yyyy-mm-dd[Thh:mm[:ss]]`, or the locale's numeric short date (its day/month/year order, any of
 * `/ - .` as one separator) with an optional `h:mm[:ss] [AM|PM]` time. Returns a day serial or null.
 * Dots only make a date in a locale that writes dates with dots, and even then need a two- or
 * four-digit year, so version numbers like "1.2.3" stay text.
 */
class LocaleDateParser {
  private readonly order: readonly DateField[];
  private readonly separator: string;

  constructor(locale: string) {
    let order: DateField[] = ['month', 'day', 'year'];
    let separator = '/';
    try {
      const parts = new Intl.DateTimeFormat(locale, {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
      }).formatToParts(new Date(Date.UTC(2006, 0, 2, 12)));
      const fields = parts
        .map((p) => p.type)
        .filter((t): t is DateField => t === 'day' || t === 'month' || t === 'year');
      if (fields.length === 3) order = fields;
      const literal = parts.find((p) => p.type === 'literal')?.value.trim();
      if (literal) separator = literal;
    } catch {
      // An unknown locale falls back to month/day/year (the invariant culture's order).
    }
    this.order = order;
    this.separator = separator;
  }

  parse(raw: string): number | null {
    const t = raw.trim();
    const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](.+))?$/.exec(t);
    if (iso) return build(Number(iso[1]), Number(iso[2]), Number(iso[3]), iso[4]);

    const local = /^(\d{1,4})([/.-])(\d{1,2})\2(\d{1,4})\.?(?:\s+(.+))?$/.exec(t);
    if (!local) return null;
    const dotted = local[2] === '.';
    if (dotted && this.separator !== '.') return null;
    const values: Record<DateField, number> = { day: 0, month: 0, year: 0 };
    const fieldTexts = [local[1]!, local[3]!, local[4]!];
    this.order.forEach((field, i) => (values[field] = Number(fieldTexts[i])));
    let year = values.year;
    const yearText = fieldTexts[this.order.indexOf('year')]!;
    if (yearText.length === 3 || (dotted && yearText.length === 1)) return null;
    if (yearText.length <= 2) year += year <= TWO_DIGIT_YEAR_PIVOT ? 2000 : 1900;
    return build(year, values.month, values.day, local[5]);
  }
}

function build(year: number, month: number, day: number, timeText: string | undefined): number | null {
  if (year < 1 || year > 9999 || month < 1 || month > 12 || day < 1 || day > daysInMonth(year, month))
    return null;
  let hour = 0;
  let minute = 0;
  let second = 0;
  if (timeText !== undefined) {
    const m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?$/.exec(timeText.trim());
    if (!m) return null;
    hour = Number(m[1]);
    minute = Number(m[2]);
    second = m[3] === undefined ? 0 : Number(m[3]);
    const meridiem = m[4]?.toLowerCase();
    if (meridiem) {
      if (hour < 1 || hour > 12) return null;
      hour = (hour % 12) + (meridiem === 'pm' ? 12 : 0);
    }
    if (hour > 23 || minute > 59 || second > 59) return null;
  }
  return DateConversion.toSerial({ year, month, day, hour, minute, second, millisecond: 0 });
}

function daysInMonth(year: number, month: number): number {
  if (month === 2) return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0 ? 29 : 28;
  return [4, 6, 9, 11].includes(month) ? 30 : 31;
}
