import { civilFromDays, DateConversion, daysFromCivil } from '../formulas/evaluation/dateConversion.ts';
import { localeInfo, safeLocale } from '../formulas/evaluation/localeInfo.ts';
import { parseNumberAny } from '../formulas/evaluation/valueParsing.ts';
import { NumberFormats, type NumberFormat } from '../styling/cellStyle.ts';
import { CellValue } from '../values/cellValue.ts';

/** What a typed cell entry means: a formula (source without `=`), or a literal plus a suggested format. */
export interface ParsedInput {
  readonly formula: string | null;
  readonly value: CellValue;
  /** A format the entry implies (a typed date) — applied only over a General format. */
  readonly suggestedFormat: NumberFormat | null;
}

/**
 * Interprets text typed into a cell (PRD §8.3) in the given locale: `'` forces text, `=` starts a
 * formula, TRUE/FALSE are booleans, then a locale number, then an ISO or slash date, else text.
 */
export const InputParser = {
  parse(text: string, locale: string): ParsedInput {
    if (!text) return literal(CellValue.empty);
    // The apostrophe is an input flag, not part of the stored value ('007, '=notaformula).
    if (text[0] === "'") return literal(CellValue.text(text.slice(1)));
    if (text[0] === '=' && text.length > 1)
      return { formula: text.slice(1), value: CellValue.empty, suggestedFormat: null };
    const upper = text.toUpperCase();
    if (upper === 'TRUE') return literal(CellValue.boolean(true));
    if (upper === 'FALSE') return literal(CellValue.boolean(false));
    const number = parseNumberAny(text, localeInfo(locale));
    if (number !== null) return literal(CellValue.number(number));
    const date = tryParseDate(text, locale);
    if (date) return { formula: null, value: CellValue.number(date.serial), suggestedFormat: date.format };
    return literal(CellValue.text(text));
  },
} as const;

function literal(value: CellValue): ParsedInput {
  return { formula: null, value, suggestedFormat: null };
}

function tryParseDate(text: string, locale: string): { serial: number; format: NumberFormat } | null {
  const trimmed = text.trim();
  if (isIsoDate(trimmed)) {
    const parts = trimmed.split('-');
    if (parts.length === 3) {
      const serial = civilSerial(Number(parts[0]), Number(parts[1]), Number(parts[2]));
      if (serial !== null) return { serial, format: NumberFormats.of('yyyy-mm-dd') };
    }
  }
  if (trimmed.includes('/')) {
    const serial = parseSlashDate(trimmed, locale);
    if (serial !== null) return { serial, format: NumberFormats.of(shortDateCode(locale)) };
  }
  return null;
}

/** yyyy-m-d with 1- or 2-digit month/day: a dash at index 4, another later, digits and dashes only. */
function isIsoDate(text: string): boolean {
  const first = text.indexOf('-');
  const last = text.lastIndexOf('-');
  return first === 4 && last !== first && /^[\d-]+$/.test(text);
}

/** Numeric d/m/y in the locale's field order; two-digit years on Calc's 1930–2029 window. */
function parseSlashDate(text: string, locale: string): number | null {
  const fields = text.split('/').map((f) => f.trim());
  if (fields.length !== 3 || fields.some((f) => !/^\d{1,4}$/.test(f))) return null;
  const order = localeInfo(locale).dateOrder;
  let year = 0;
  let month = 0;
  let day = 0;
  order.forEach((part, i) => {
    const f = fields[i]!;
    const n = Number(f);
    if (part === 'year') year = f.length <= 2 ? (n <= 29 ? 2000 + n : 1900 + n) : n;
    else if (part === 'month') month = n;
    else day = n;
  });
  return civilSerial(year, month, day);
}

/** The serial of a real calendar date, or null (31 February, month 13, year 0). */
function civilSerial(year: number, month: number, day: number): number | null {
  if (year < 1 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return null;
  const back = civilFromDays(daysFromCivil(year, month, day));
  if (back.year !== year || back.month !== month || back.day !== day) return null;
  return DateConversion.toSerial({ year, month, day, hour: 0, minute: 0, second: 0, millisecond: 0 });
}

const shortDateCodes = new Map<string, string>();

/** The locale's numeric short-date pattern as a format code (`m/d/yyyy`, `dd.mm.yyyy`). */
export function shortDateCode(locale: string): string {
  let code = shortDateCodes.get(locale);
  if (code !== undefined) return code;
  const fmt = new Intl.DateTimeFormat(safeLocale(locale), {
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    timeZone: 'UTC',
  });
  // 2 January: a padded field reads "01"/"02", an unpadded one "1"/"2".
  code = '';
  for (const p of fmt.formatToParts(new Date(Date.UTC(2026, 0, 2)))) {
    if (p.type === 'year') code += 'yyyy';
    else if (p.type === 'month') code += p.value.length > 1 ? 'mm' : 'm';
    else if (p.type === 'day') code += p.value.length > 1 ? 'dd' : 'd';
    else if (p.type === 'literal') code += /[a-zA-Z]/.test(p.value) ? `"${p.value}"` : p.value;
  }
  if (!code.includes('yyyy')) code = 'yyyy-mm-dd';
  shortDateCodes.set(locale, code);
  return code;
}
