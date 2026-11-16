import { DateConversion, type CivilDateTime } from '../../formulas/evaluation/dateConversion.ts';
import { NumberFormats, type NumberFormat } from '../../styling/cellStyle.ts';
import { NumberFormatter } from '../../styling/numberFormatter.ts';
import { CellError } from '../../values/cellError.ts';
import { CellValue, roundTripNumber } from '../../values/cellValue.ts';
import type { XmlElement } from '../../xml/index.ts';
import { parseInvariantFloat } from './mapping/odsCellStyles.ts';
import { paragraphText } from './odsContentShared.ts';
import { Calcext, Office } from './odsXmlNames.ts';

/**
 * How a cell's value is stored in ODF: `office:value-type` and the typed value attribute beside it
 * (`office:value`, `office:date-value`, `office:time-value`, `office:boolean-value`,
 * `office:string-value`), plus LibreOffice's `calcext:value-type="error"`. Reading and writing live
 * together so the two stay each other's inverse.
 */

/** A fixed locale for the cached display text written beside each value, whatever the system's. */
const INVARIANT_LOCALE = 'en-US';

// ---- reading ----------------------------------------------------------------

/** The value a `table:table-cell` holds; `paragraphs` are its `text:p` children (its display). */
export function readCellValue(element: XmlElement, paragraphs: readonly XmlElement[]): CellValue {
  if (element.attr(Calcext('value-type')) === 'error') {
    // LibreOffice writes an error result as an empty string marked as an error, the token as display.
    // A token we have no code for (Calc's "Err:502") stays the text it shows.
    const shown = paragraphs.map(paragraphText).join('\n').trim();
    const code = CellError.codeFromToken(shown);
    return code === null
      ? CellValue.text(shown)
      : CellValue.error(new CellError(code, `The file saved this result as ${shown}.`));
  }
  switch (element.attr(Office('value-type'))) {
    case 'float':
    case 'percentage':
    case 'currency': {
      const n = parseDouble(element.attr(Office('value')));
      return n === null ? CellValue.empty : CellValue.number(n);
    }
    case 'date': {
      const serial = parseDateValue(element.attr(Office('date-value')));
      return serial === null ? CellValue.empty : CellValue.number(serial);
    }
    case 'time': {
      const days = parseDuration(element.attr(Office('time-value')));
      return days === null ? CellValue.empty : CellValue.number(days);
    }
    case 'boolean': {
      // xsd:boolean is "true"/"false" or "1"/"0".
      const b = element.attr(Office('boolean-value'))?.trim();
      return CellValue.boolean(b === 'true' || b === '1');
    }
    case 'string': {
      // office:string-value, when present, is the value; the paragraphs are only its display.
      const value = element.attr(Office('string-value'));
      return CellValue.text(value ?? paragraphs.map(paragraphText).join('\n'));
    }
    case null:
    default: {
      // No value type, but text: a string cell from a writer that left the type out (ODF 19.385
      // makes it optional), which LibreOffice reads as text too.
      const text = paragraphs.map(paragraphText).join('\n');
      return text.length > 0 ? CellValue.text(text) : CellValue.empty;
    }
  }
}

export function parseDouble(text: string | null): number | null {
  switch (text) {
    case 'INF':
      return Infinity;
    case '-INF':
      return -Infinity;
    case 'NaN':
      return NaN;
    case null:
      break;
  }
  return parseInvariantFloat(text);
}

const ISO_DATE =
  /^\s*(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?\s*(?:Z|[+-]\d{2}:?\d{2})?\s*$/;

/** An ODF `office:date-value` (`yyyy-MM-dd` or `yyyy-MM-ddTHH:mm:ss[.fff]`) → a day serial. */
function parseDateValue(text: string | null): number | null {
  if (text === null) return null;
  const m = ISO_DATE.exec(text);
  if (!m) return null;
  const year = Number(m[1]);
  const month = Number(m[2]);
  const day = Number(m[3]);
  const hour = m[4] === undefined ? 0 : Number(m[4]);
  const minute = m[5] === undefined ? 0 : Number(m[5]);
  const second = m[6] === undefined ? 0 : Number(m[6]);
  const millisecond = m[7] === undefined ? 0 : Math.round(Number('0.' + m[7]) * 1000);
  if (month < 1 || month > 12 || day < 1 || day > daysInMonth(year, month)) return null;
  if (hour > 23 || minute > 59 || second > 59) return null;
  return DateConversion.toSerial({ year, month, day, hour, minute, second, millisecond });
}

function daysInMonth(year: number, month: number): number {
  if (month === 2) return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0 ? 29 : 28;
  return [4, 6, 9, 11].includes(month) ? 30 : 31;
}

const DURATION =
  /^\s*(-)?P(?:(\d+)Y)?(?:(\d+)M)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d*)?|\.\d+)S)?)?\s*$/;

/**
 * `XmlConvert.ToTimeSpan(text).TotalDays`: an `xsd:duration` → days, with .NET's conventions of a
 * 365-day year and a 30-day month. Null when malformed (including a bare `P` or a dangling `T`).
 */
export function parseDuration(text: string | null): number | null {
  if (!text) return null;
  const m = DURATION.exec(text);
  if (!m) return null;
  const hasDate = m[2] !== undefined || m[3] !== undefined || m[4] !== undefined;
  const hasTime = m[5] !== undefined || m[6] !== undefined || m[7] !== undefined;
  if (!hasDate && !hasTime) return null;
  if (text.includes('T') && !hasTime) return null;

  const n = (s: string | undefined): number => (s === undefined ? 0 : Number(s));
  const seconds = ((n(m[2]) * 365 + n(m[3]) * 30 + n(m[4])) * 24 + n(m[5])) * 3600 + n(m[6]) * 60 + n(m[7]);
  const ticks = Math.round(seconds * 10_000_000);
  const days = ticks / 864_000_000_000;
  return m[1] ? -days : days;
}

// ---- writing ----------------------------------------------------------------

/**
 * Type a number, boolean or error value onto `element` as ODF stores it, and return the display
 * lines to write beside it. A number's type follows its number format: a date format makes it an
 * `office:date-value`, a time format an `xsd:duration`, and so on. Text values are the writer's
 * own business (rich text, line breaks).
 */
export function writeCellValue(
  element: XmlElement,
  value: Exclude<CellValue, { kind: 'text' } | { kind: 'empty' }>,
  format: NumberFormat,
): string[] {
  switch (value.kind) {
    case 'number': {
      const kind = NumberFormats.kind(format);
      if (kind === 'time') {
        // A time format stores the serial as an ISO-8601 duration so the whole value (days *and*
        // time-of-day fraction) survives, including elapsed times past 24h.
        element.setAttr(Office('value-type'), 'time');
        element.setAttr(Office('time-value'), durationText(value.value));
      } else if (kind === 'date' || kind === 'dateTime') {
        element.setAttr(Office('value-type'), 'date');
        element.setAttr(Office('date-value'), dateValueText(value.value));
      } else if (kind === 'currency') {
        element.setAttr(Office('value-type'), 'currency');
        element.setAttr(Office('currency'), currencyCode(format.code));
        element.setAttr(Office('value'), floatValueText(value.value));
      } else if (kind === 'percent') {
        element.setAttr(Office('value-type'), 'percentage');
        element.setAttr(Office('value'), floatValueText(value.value));
      } else {
        element.setAttr(Office('value-type'), 'float');
        element.setAttr(Office('value'), floatValueText(value.value));
      }
      return [NumberFormatter.format(value, format.code, INVARIANT_LOCALE)];
    }
    case 'boolean':
      element.setAttr(Office('value-type'), 'boolean');
      element.setAttr(Office('boolean-value'), value.value ? 'true' : 'false');
      return [value.value ? 'TRUE' : 'FALSE'];
    case 'error':
      // As LibreOffice writes it: an empty string value marked as an error, the token as display.
      element.setAttr(Office('value-type'), 'string');
      element.setAttr(Office('string-value'), '');
      element.setAttr(Calcext('value-type'), 'error');
      return [value.error.token];
  }
}

const CURRENCY_BY_SYMBOL: Readonly<Record<string, string>> = { $: 'USD', '€': 'EUR', '£': 'GBP', '¥': 'JPY' };

/**
 * The ISO 4217 code for `office:currency`, read from the format's symbol: `[$€-407]`, `[$USD]`,
 * `"$"` or a bare `€`. Null (the attribute is optional) when the code names no currency we know.
 */
function currencyCode(code: string): string | null {
  const tagged = /\[\$([^\]-]+)/.exec(code)?.[1];
  const symbol = tagged ?? /[$€£¥]/.exec(code)?.[0];
  if (symbol === undefined) return null;
  if (/^[A-Z]{3}$/.test(symbol)) return symbol;
  return CURRENCY_BY_SYMBOL[symbol] ?? null;
}

/**
 * An ODF `office:value` (`xsd:double`). Non-finite values spell `INF`/`-INF`/`NaN`; "Infinity" is
 * rejected by conformant readers.
 */
function floatValueText(n: number): string {
  if (Number.isNaN(n)) return 'NaN';
  if (n === Infinity) return 'INF';
  if (n === -Infinity) return '-INF';
  return roundTripNumber(n);
}

/**
 * An ODF `office:date-value`. Pure dates stay `yyyy-MM-dd`; values with a time-of-day fraction emit
 * the full ISO-8601 timestamp. Out-of-range serials fall back to a plain date string.
 */
function dateValueText(serial: number): string {
  const when = DateConversion.tryToDateTime(serial, false);
  if (!when) return datePart(DateConversion.dateFromSerial(serial));
  const midnight = when.hour === 0 && when.minute === 0 && when.second === 0 && when.millisecond === 0;
  if (midnight) return datePart(when);
  // Sub-second time is part of the value: 12:00:00.5 must not come back as 12:00:00.
  const fraction = when.millisecond === 0 ? '' : '.' + pad(when.millisecond, 3).replace(/0+$/, '');
  return `${datePart(when)}T${pad(when.hour, 2)}:${pad(when.minute, 2)}:${pad(when.second, 2)}${fraction}`;
}

function datePart(d: CivilDateTime): string {
  return `${pad(d.year, 4)}-${pad(d.month, 2)}-${pad(d.day, 2)}`;
}

function pad(n: number, width: number): string {
  return String(n).padStart(width, '0');
}

const TICKS_PER_DAY = 864_000_000_000;
const TICKS_PER_SECOND = 10_000_000;

/**
 * `XmlConvert.ToString(TimeSpan.FromDays(days))`: an `xsd:duration` such as `P1DT6H`, `PT12H`,
 * `PT0.5S`, or `PT0S`, at tick (100 ns) precision with a trimmed seven-digit fraction.
 */
export function durationText(days: number): string {
  if (!Number.isFinite(days)) return 'PT0S';
  let ticks = Math.round(days * TICKS_PER_DAY);
  let sign = '';
  if (ticks < 0) {
    sign = '-';
    ticks = -ticks;
  }

  const d = Math.floor(ticks / TICKS_PER_DAY);
  let rest = ticks - d * TICKS_PER_DAY;
  const h = Math.floor(rest / (3600 * TICKS_PER_SECOND));
  rest -= h * 3600 * TICKS_PER_SECOND;
  const m = Math.floor(rest / (60 * TICKS_PER_SECOND));
  rest -= m * 60 * TICKS_PER_SECOND;
  const s = Math.floor(rest / TICKS_PER_SECOND);
  const fraction = rest - s * TICKS_PER_SECOND;

  let text = sign + 'P';
  if (d !== 0) text += `${d}D`;
  if (h !== 0 || m !== 0 || s !== 0 || fraction !== 0) {
    text += 'T';
    if (h !== 0) text += `${h}H`;
    if (m !== 0) text += `${m}M`;
    if (s !== 0 || fraction !== 0) {
      text += String(s);
      if (fraction !== 0) text += '.' + String(fraction).padStart(7, '0').replace(/0+$/, '');
      text += 'S';
    }
  } else if (d === 0) {
    text += 'T0S';
  }
  return text;
}
