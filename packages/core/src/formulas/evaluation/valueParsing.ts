import { DateConversion, daysFromCivil } from './dateConversion.ts';
import { localeInfo, type LocaleInfo } from './localeInfo.ts';

interface NumberSymbols {
  readonly decimal: string;
  readonly group: string;
  readonly currencySymbol: string;
}

/** The .NET invariant culture's number symbols. */
export const INVARIANT_SYMBOLS: NumberSymbols = { decimal: '.', group: ',', currencySymbol: '¤' };

const escapeRegex = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const isSpace = (s: string): boolean => /^\s$/u.test(s) || s === ' ';

/**
 * `double.TryParse(text, NumberStyles.Any, culture)`: surrounding whitespace, a leading or trailing
 * sign, parentheses for negatives, the culture's currency symbol, group separators in the integer
 * part, one decimal separator, and an exponent. Null when the text is not such a number.
 */
export function parseNumberAny(text: string, symbols: NumberSymbols): number | null {
  let t = text.trim();
  if (t === '') return null;
  let negative = false;
  const paren = /^\((.*)\)$/s.exec(t);
  if (paren) {
    negative = true;
    t = paren[1]!.trim();
  }
  const takeCurrency = (): void => {
    const cs = symbols.currencySymbol;
    if (cs && t.startsWith(cs)) t = t.slice(cs.length).trim();
    else if (cs && t.endsWith(cs)) t = t.slice(0, -cs.length).trim();
  };
  takeCurrency();
  const sign = /^[+-]/.exec(t) ?? /[+-]$/.exec(t);
  if (sign) {
    if (paren) return null;
    negative = sign[0] === '-';
    t = sign.index === 0 ? t.slice(1).trim() : t.slice(0, -1).trim();
    takeCurrency();
  }
  const group = isSpace(symbols.group) ? `[${escapeRegex(symbols.group)}\\s]` : escapeRegex(symbols.group);
  const decimal = escapeRegex(symbols.decimal);
  const re = new RegExp(`^(\\d(?:\\d|${group})*)?(?:${decimal}(\\d*))?(?:[eE]([+-]?\\d+))?$`, 'u');
  const m = re.exec(t);
  if (!m || (m[1] === undefined && !m[2])) return null;
  const integer = (m[1] ?? '').replace(new RegExp(group, 'gu'), '');
  const n = Number(`${integer || '0'}.${m[2] ?? ''}e${m[3] ?? '0'}`);
  if (!Number.isFinite(n)) return null;
  return negative ? -n : n;
}

// ---- dates -----------------------------------------------------------------

const TIME = /^(\d{1,2}):(\d{2})(?::(\d{2})(?:[.,](\d+))?)?\s*(?:([AaPp])\.?\s*[Mm]?\.?)?$/;

/** "13:45", "1:45 PM", "9:05:30" → fraction of a day; null when it is not a time. */
function parseTime(text: string, info: LocaleInfo): number | null {
  let t = text.trim();
  let meridiem: 'am' | 'pm' | null = null;
  const lower = t.toLocaleLowerCase();
  for (const [designator, which] of [
    [info.am, 'am'],
    [info.pm, 'pm'],
  ] as const) {
    const d = designator.toLocaleLowerCase();
    if (d && d !== 'am' && d !== 'pm' && lower.endsWith(d)) {
      meridiem = which;
      t = t.slice(0, t.length - d.length).trim();
    }
  }
  const m = TIME.exec(t);
  if (!m) return null;
  let hour = Number(m[1]);
  const minute = Number(m[2]);
  const second = m[3] ? Number(m[3]) : 0;
  const ms = m[4] ? Math.round(Number(`0.${m[4]}`) * 1000) : 0;
  if (m[5]) meridiem = m[5].toLowerCase() === 'a' ? 'am' : 'pm';
  if (meridiem) {
    if (hour < 1 || hour > 12) return null;
    hour = (hour % 12) + (meridiem === 'pm' ? 12 : 0);
  }
  if (hour > 23 || minute > 59 || second > 59) return null;
  return (((hour * 60 + minute) * 60 + second) * 1000 + ms) / 86_400_000;
}

function monthFromName(word: string, info: LocaleInfo): number | null {
  const w = word.replace(/\.$/, '').toLocaleLowerCase();
  if (w.length < 3) return null;
  const lists = [info.months, info.monthsShort, ENGLISH_MONTHS];
  for (const list of lists) {
    const i = list.findIndex((n) => {
      const name = n.replace(/\.$/, '').toLocaleLowerCase();
      return name === w || (w.length >= 3 && name.startsWith(w) && list !== info.months);
    });
    if (i >= 0) return i + 1;
  }
  return null;
}

const ENGLISH_MONTHS = [
  'january',
  'february',
  'march',
  'april',
  'may',
  'june',
  'july',
  'august',
  'september',
  'october',
  'november',
  'december',
];

function fullYear(text: string): number {
  const y = Number(text);
  if (text.length <= 2) return y + (y <= 29 ? 2000 : 1900);
  return y;
}

const LONGEST_MONTH = 31;

/** The serial of a real calendar day in years 1–9999, or null. Never throws. */
function validDate(year: number, month: number, day: number): number | null {
  if (year < 1 || year > 9999 || month < 1 || month > 12 || day < 1 || day > LONGEST_MONTH) return null;
  // Check the day against the month first: DATE-style roll-over past 9999-12-31 throws.
  // (daysFromCivil reads month 13 as January of the next year.)
  const daysInMonth = daysFromCivil(year, month + 1, 1) - daysFromCivil(year, month, 1);
  if (day > daysInMonth) return null;
  return DateConversion.dateToSerial(year, month, day);
}

function parseDatePart(text: string, info: LocaleInfo): number | null {
  const t = text.trim().replace(/,/g, ' ').replace(/\s+/g, ' ');
  // ISO yyyy-mm-dd (or yyyy/mm/dd).
  let m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(t);
  if (m) return validDate(Number(m[1]), Number(m[2]), Number(m[3]));
  // Numeric in the locale's field order: 6/7/2026, 07.06.2026, 6/7 (current year not modelled).
  m = /^(\d{1,4})[-/. ](\d{1,2})[-/. ](\d{1,4})$/.exec(t);
  if (m) {
    const fields = { day: 0, month: 0, year: 0 };
    let yearText = '';
    info.dateOrder.forEach((f, i) => {
      const v = m![i + 1]!;
      if (f === 'year') yearText = v;
      else fields[f] = Number(v);
    });
    return validDate(fullYear(yearText), fields.month, fields.day);
  }
  // Month names: "June 7 2026", "7 June 2026", "7-Jun-2026", "2026 June 7".
  const words = t.split(/[\s\-/.]+/).filter((w) => w !== '');
  if (words.length === 3) {
    const monthIndex = words.findIndex((w) => /\p{L}/u.test(w));
    const month = monthIndex >= 0 ? monthFromName(words[monthIndex]!, info) : null;
    if (month === null) return null;
    const nums = words.filter((_, i) => i !== monthIndex);
    if (!nums.every((w) => /^\d+$/.test(w))) return null;
    const [a, b] = nums as [string, string];
    // The year is the 4-digit one, else the last.
    const yearFirst = a.length >= 3;
    const year = fullYear(yearFirst ? a : b);
    const day = Number(yearFirst ? b : a);
    return validDate(year, month, day);
  }
  return null;
}

/**
 * A modest stand-in for `DateTime.TryParse(text, culture)`: ISO dates, numeric dates in the locale's
 * field order, dates with month names (locale or English), each optionally followed by a time, or a
 * time alone. Returns the serial (a time alone is just the day fraction), or null.
 */
export function parseDateTime(text: string, locale: string): number | null {
  const info = localeInfo(locale);
  const t = text.trim();
  if (t === '') return null;
  const time = parseTime(t, info);
  if (time !== null) return time;
  const iso = /^(\d{4}-\d{1,2}-\d{1,2})[T ](.+)$/.exec(t);
  if (iso) {
    const d = parseDatePart(iso[1]!, info);
    const tm = parseTime(iso[2]!.replace(/Z$/, ''), info);
    return d !== null && tm !== null ? d + tm : null;
  }
  const whole = parseDatePart(t, info);
  if (whole !== null) return whole;
  // Date followed by a time: split at the first space before a digit:digit run.
  const split = /^(.*?)\s+(\d{1,2}:\d{2}.*)$/.exec(t);
  if (split) {
    const d = parseDatePart(split[1]!, info);
    const tm = parseTime(split[2]!, info);
    if (d !== null && tm !== null) return d + tm;
  }
  return null;
}

/**
 * Parse a value for VALUE() and lookups, using the active locale: a trailing `%` divides by 100,
 * then a locale or invariant number (NumberStyles.Any), then a leading `$` or the locale's currency
 * symbol stripped, then a date/time. Null when nothing matches.
 */
export function tryParseNumberOrDate(input: string, locale: string): number | null {
  const info = localeInfo(locale);
  const text = input.trim();
  if (text.endsWith('%')) {
    const body = text.slice(0, -1).trim();
    const n = parseNumberAny(body, info) ?? parseNumberAny(body, INVARIANT_SYMBOLS);
    if (n !== null) return n / 100;
  }
  const direct = parseNumberAny(text, info) ?? parseNumberAny(text, INVARIANT_SYMBOLS);
  if (direct !== null) return direct;

  let stripped: string | null = null;
  if (text.startsWith('$')) stripped = text.slice(1).trimStart();
  else if (info.currencySymbol && text.startsWith(info.currencySymbol)) {
    stripped = text.slice(info.currencySymbol.length).trimStart();
  }
  if (stripped !== null) {
    const n = parseNumberAny(stripped, info) ?? parseNumberAny(stripped, INVARIANT_SYMBOLS);
    if (n !== null) return n;
  }
  return parseDateTime(text, locale);
}
