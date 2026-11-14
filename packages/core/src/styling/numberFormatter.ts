import { DateConversion, type CivilDateTime } from '../formulas/evaluation/dateConversion.ts';
import { fixedDigits } from '../formulas/evaluation/decimalDigits.ts';
import { localeInfo, type LocaleInfo } from '../formulas/evaluation/localeInfo.ts';
import type { CellValue } from '../values/cellValue.ts';

// ---- tokens -----------------------------------------------------------------

export type FormatToken =
  | { readonly t: 'literal'; readonly text: string }
  | { readonly t: 'digit'; readonly ch: '0' | '#' | '?' }
  | { readonly t: 'point' }
  | { readonly t: 'comma' }
  | { readonly t: 'percent' }
  | { readonly t: 'exponent'; readonly showPlus: boolean; readonly upper: boolean }
  | { readonly t: 'text' }
  | { readonly t: 'general' }
  | { readonly t: 'date'; readonly field: DateField; readonly length: number }
  | { readonly t: 'ampm'; readonly long: boolean; readonly lower: boolean }
  | { readonly t: 'elapsed'; readonly unit: 'h' | 'm' | 's'; readonly length: number };

export type DateField = 'y' | 'M' | 'd' | 'h' | 'm' | 's' | 'f';

export interface Condition {
  readonly op: '<' | '<=' | '>' | '>=' | '=' | '<>';
  readonly value: number;
}

export interface Section {
  readonly tokens: readonly FormatToken[];
  readonly condition: Condition | null;
  readonly isDate: boolean;
}

/** Split a code on `;` outside quotes, brackets and escapes. */
export function splitSections(code: string): string[] {
  const sections: string[] = [];
  let current = '';
  for (let i = 0; i < code.length; i++) {
    const c = code[i]!;
    if (c === '"') {
      const end = code.indexOf('"', i + 1);
      const stop = end < 0 ? code.length : end + 1;
      current += code.slice(i, stop);
      i = stop - 1;
    } else if (c === '[') {
      const end = code.indexOf(']', i + 1);
      const stop = end < 0 ? code.length : end + 1;
      current += code.slice(i, stop);
      i = stop - 1;
    } else if (c === '\\' || c === '_' || c === '*') {
      current += code.slice(i, i + 2);
      i++;
    } else if (c === ';') {
      sections.push(current);
      current = '';
    } else current += c;
  }
  sections.push(current);
  return sections;
}

const runLength = (s: string, i: number, c: string): number => {
  let n = 0;
  while (i + n < s.length && s[i + n]!.toLowerCase() === c) n++;
  return n;
};

/** Parse a bracket's content: a colour/locale/condition/elapsed token. */
function bracket(inside: string, tokens: FormatToken[]): Condition | null {
  const elapsed = /^(h+|m+|s+)$/i.exec(inside);
  if (elapsed) {
    const unit = inside[0]!.toLowerCase() as 'h' | 'm' | 's';
    tokens.push({ t: 'elapsed', unit, length: inside.length });
    return null;
  }
  const condition = /^(<=|>=|<>|<|>|=)\s*(-?\d+(?:\.\d+)?)$/.exec(inside);
  if (condition) return { op: condition[1] as Condition['op'], value: Number(condition[2]) };
  if (inside.startsWith('$')) {
    // [$€-407] / [$USD] / [$-409]: the symbol before the dash is shown; the locale tag is not.
    const dash = inside.indexOf('-', 1);
    const symbol = inside.slice(1, dash < 0 ? inside.length : dash);
    if (symbol) tokens.push({ t: 'literal', text: symbol });
  }
  // Colours ([Red], [Color 3]) and anything else unknown are display hints: ignored.
  return null;
}

/** Tokenize one section of a format code (the `.ods` writer maps date and time tokens to ODF). */
export function tokenize(code: string): Section {
  const tokens: FormatToken[] = [];
  let condition: Condition | null = null;
  const lower = code.toLowerCase();
  for (let i = 0; i < code.length; i++) {
    const c = code[i]!;
    const lc = lower[i]!;
    if (c === '"') {
      const end = code.indexOf('"', i + 1);
      const stop = end < 0 ? code.length : end;
      tokens.push({ t: 'literal', text: code.slice(i + 1, stop) });
      i = stop;
    } else if (c === '\\') {
      if (i + 1 < code.length) tokens.push({ t: 'literal', text: code[i + 1]! });
      i++;
    } else if (c === '_') {
      tokens.push({ t: 'literal', text: ' ' }); // space the width of the next char
      i++;
    } else if (c === '*') {
      i++; // repeat-to-fill: no fixed width to fill here
    } else if (c === '[') {
      const end = code.indexOf(']', i + 1);
      const stop = end < 0 ? code.length : end;
      condition = bracket(code.slice(i + 1, stop), tokens) ?? condition;
      i = stop;
    } else if (c === '0' || c === '#' || c === '?') {
      tokens.push({ t: 'digit', ch: c });
    } else if (c === '.') {
      // After a seconds token a point starts fractional seconds ("ss.00").
      const prev = tokens[tokens.length - 1];
      if (
        prev &&
        (prev.t === 'date' ? prev.field === 's' : prev.t === 'elapsed' && prev.unit === 's') &&
        code[i + 1] === '0'
      ) {
        const n = runLength(code, i + 1, '0');
        tokens.push({ t: 'date', field: 'f', length: n });
        i += n;
      } else tokens.push({ t: 'point' });
    } else if (c === ',') {
      tokens.push({ t: 'comma' });
    } else if (c === '%') {
      tokens.push({ t: 'percent' });
    } else if ((c === 'E' || c === 'e') && (code[i + 1] === '+' || code[i + 1] === '-')) {
      tokens.push({ t: 'exponent', showPlus: code[i + 1] === '+', upper: c === 'E' });
      i++;
    } else if (lower.startsWith('general', i)) {
      tokens.push({ t: 'general' });
      i += 6;
    } else if (c === '@') {
      tokens.push({ t: 'text' });
    } else if (lower.startsWith('am/pm', i)) {
      tokens.push({ t: 'ampm', long: true, lower: code.slice(i, i + 5) === 'am/pm' });
      i += 4;
    } else if (lower.startsWith('a/p', i)) {
      tokens.push({ t: 'ampm', long: false, lower: code.slice(i, i + 3) === 'a/p' });
      i += 2;
    } else if (lc === 'y' || lc === 'd' || lc === 'h' || lc === 's' || lc === 'm') {
      const n = runLength(lower, i, lc);
      const field: DateField =
        lc === 'y' ? 'y' : lc === 'd' ? 'd' : lc === 'h' ? 'h' : lc === 's' ? 's' : 'M';
      tokens.push({ t: 'date', field, length: n });
      i += n - 1;
    } else {
      tokens.push({ t: 'literal', text: c });
    }
  }
  resolveMinutes(tokens);
  const isDate = tokens.some((tk) => tk.t === 'date' || tk.t === 'elapsed' || tk.t === 'ampm');
  return { tokens, condition, isDate };
}

/**
 * `m`/`mm` is minutes when the nearest date token before it is an hour, or the nearest after it is
 * seconds (separators and literals in between are skipped); otherwise it is the month.
 */
function resolveMinutes(tokens: FormatToken[]): void {
  const isTime = (tk: FormatToken | undefined, field: 'h' | 's'): boolean =>
    !!tk && ((tk.t === 'date' && tk.field === field) || (tk.t === 'elapsed' && tk.unit === field));
  const nearest = (from: number, step: number): FormatToken | undefined => {
    for (let j = from; j >= 0 && j < tokens.length; j += step) {
      const tk = tokens[j]!;
      if (tk.t === 'date' || tk.t === 'elapsed' || tk.t === 'ampm') return tk;
    }
    return undefined;
  };
  for (let i = 0; i < tokens.length; i++) {
    const tk = tokens[i]!;
    if (tk.t !== 'date' || tk.field !== 'M' || tk.length > 2) continue;
    if (isTime(nearest(i - 1, -1), 'h') || isTime(nearest(i + 1, 1), 's')) {
      tokens[i] = { t: 'date', field: 'm', length: tk.length };
    }
  }
}

// ---- sections ---------------------------------------------------------------

function matches(c: Condition, n: number): boolean {
  switch (c.op) {
    case '<':
      return n < c.value;
    case '<=':
      return n <= c.value;
    case '>':
      return n > c.value;
    case '>=':
      return n >= c.value;
    case '=':
      return n === c.value;
    case '<>':
      return n !== c.value;
  }
}

/** Pick the section for `n`; `signed` says whether the section shows the sign itself. */
function pickSection(sections: readonly Section[], n: number): { section: Section; signed: boolean } {
  const numeric = sections.slice(0, 3);
  if (numeric.some((s) => s.condition)) {
    for (const s of numeric) if (s.condition && matches(s.condition, n)) return { section: s, signed: false };
    const rest = numeric.filter((s) => !s.condition);
    return { section: rest[0] ?? numeric[numeric.length - 1]!, signed: rest.length === 0 };
  }
  if (numeric.length === 1) return { section: numeric[0]!, signed: true };
  if (n < 0) return { section: numeric[1]!, signed: false };
  if (n === 0 && numeric.length >= 3) return { section: numeric[2]!, signed: false };
  return { section: numeric[0]!, signed: true };
}

// ---- numbers ----------------------------------------------------------------

const placeholder = (tokens: readonly FormatToken[], i: number): string => (tokens[i] as { ch: string }).ch;

/** Fill digit placeholders right to left from `digits`; the leftmost one takes any overflow. */
function assignRightToLeft(
  digits: string,
  idxs: readonly number[],
  tokens: readonly FormatToken[],
  out: Map<number, string>,
): void {
  const n = idxs.length;
  idxs.forEach((i, k) => {
    const fromRight = n - 1 - k;
    const ch = placeholder(tokens, i);
    let s: string;
    if (fromRight < digits.length) s = digits[digits.length - 1 - fromRight]!;
    else s = ch === '0' ? '0' : ch === '?' ? ' ' : '';
    if (k === 0 && digits.length > n) s = digits.slice(0, digits.length - n + 1);
    out.set(i, s);
  });
}

const groupDigits = (digits: string, sep: string): string => digits.replace(/\B(?=(\d{3})+(?!\d))/g, sep);

function formatNumberSection(
  abs: number,
  tokens: readonly FormatToken[],
  info: LocaleInfo,
): { text: string; nonZero: boolean } {
  const pointAt = tokens.findIndex((tk) => tk.t === 'point');
  const expAt = tokens.findIndex((tk) => tk.t === 'exponent');
  const intEnd = pointAt >= 0 ? pointAt : expAt >= 0 ? expAt : tokens.length;
  const fracEnd = expAt >= 0 ? expAt : tokens.length;

  const intDigits: number[] = [];
  const fracDigits: number[] = [];
  const expDigits: number[] = [];
  tokens.forEach((tk, i) => {
    if (tk.t !== 'digit') return;
    if (i < intEnd) intDigits.push(i);
    else if (i < fracEnd) fracDigits.push(i);
    else expDigits.push(i);
  });

  // Commas: between integer placeholders → grouping; right after the last one → scale by 1000 each.
  let grouping = false;
  let scale = 0;
  const silentCommas = new Set<number>();
  const firstInt = intDigits[0] ?? -1;
  const lastInt = intDigits[intDigits.length - 1] ?? -1;
  const lastFrac = fracDigits[fracDigits.length - 1] ?? -1;
  tokens.forEach((tk, i) => {
    if (tk.t !== 'comma') return;
    if (firstInt >= 0 && i > firstInt && i < lastInt) {
      grouping = true;
      silentCommas.add(i);
    } else if (lastInt >= 0 && i > lastInt && i < intEnd) {
      let j = lastInt + 1;
      while (j < i && tokens[j]!.t === 'comma') j++;
      if (j === i) {
        scale++;
        silentCommas.add(i);
      }
    } else if (lastFrac >= 0 && i > lastFrac && i < fracEnd) {
      // `0.0,,` — a comma run right after the last fraction placeholder scales too.
      let j = lastFrac + 1;
      while (j < i && tokens[j]!.t === 'comma') j++;
      if (j === i) {
        scale++;
        silentCommas.add(i);
      }
    }
  });

  const percents = tokens.filter((tk) => tk.t === 'percent').length;
  let value = (abs * Math.pow(100, percents)) / Math.pow(1000, scale);
  const fracCount = fracDigits.length;

  let exponent = 0;
  if (expAt >= 0 && value !== 0) {
    const width = Math.max(1, intDigits.length);
    exponent = Math.floor(Math.log10(value)) - (width - 1);
    let mantissa = value / Math.pow(10, exponent);
    const rounded = fixedDigits(mantissa, fracCount);
    if (rounded.integer.length > width) {
      exponent++;
      mantissa = value / Math.pow(10, exponent);
    }
    value = mantissa;
  }

  const { integer, fraction } = fixedDigits(value, fracCount);
  const intText = integer === '0' ? '' : integer;
  let nonZero = /[1-9]/.test(intText + fraction);

  // Integer placeholders, right to left; the leftmost takes the overflow.
  const intOut = new Map<number, string>();
  if (grouping) {
    const firstZero = intDigits.findIndex((i) => placeholder(tokens, i) === '0');
    const minDigits = firstZero < 0 ? 0 : intDigits.length - firstZero;
    const digits = intText.padStart(minDigits, '0');
    const spaces = intDigits
      .slice(0, Math.max(0, intDigits.length - digits.length))
      .filter((i) => placeholder(tokens, i) === '?').length;
    const grouped = ' '.repeat(spaces) + groupDigits(digits, info.group);
    intDigits.forEach((i, k) => intOut.set(i, k === 0 ? grouped : ''));
  } else {
    assignRightToLeft(intText, intDigits, tokens, intOut);
  }

  // Fraction placeholders: trailing '#' zeros drop, '?' zeros become spaces.
  const fracOut = new Map<number, string>();
  let trailing = true;
  for (let k = fracCount - 1; k >= 0; k--) {
    const i = fracDigits[k]!;
    const ch = placeholder(tokens, i);
    const d = fraction[k]!;
    if (trailing && d === '0' && ch !== '0') fracOut.set(i, ch === '?' ? ' ' : '');
    else {
      trailing = false;
      fracOut.set(i, d);
    }
  }
  const anyFraction = [...fracOut.values()].some((s) => s !== '');

  // Exponent digits.
  const expOut = new Map<number, string>();
  if (expAt >= 0) assignRightToLeft(String(Math.abs(exponent)).replace(/^0$/, ''), expDigits, tokens, expOut);

  let text = '';
  tokens.forEach((tk, i) => {
    switch (tk.t) {
      case 'digit':
        text += intOut.get(i) ?? fracOut.get(i) ?? expOut.get(i) ?? '';
        break;
      case 'point':
        if (i !== pointAt) break;
        // No integer placeholder: the integer part still shows when non-zero (".00" of 3.5 → "3.50").
        if (intDigits.length === 0) text += intText;
        if (anyFraction) text += info.decimal;
        break;
      case 'comma':
        if (!silentCommas.has(i)) text += ',';
        break;
      case 'percent':
        text += '%';
        break;
      case 'exponent':
        if (intDigits.length === 0 && pointAt < 0) text += intText;
        text += (tk.upper ? 'E' : 'e') + (exponent < 0 ? '-' : tk.showPlus ? '+' : '');
        break;
      case 'literal':
        text += tk.text;
        break;
      case 'general':
        text += NumberFormatter.formatGeneral(abs, info.locale);
        break;
      default:
        break; // '@' and stray date tokens do not apply to numbers
    }
  });
  if (tokens.some((tk) => tk.t === 'general')) nonZero = abs !== 0;
  else if (intDigits.length === 0 && pointAt < 0 && expAt < 0 && fracDigits.length === 0) nonZero = false;
  return { text, nonZero };
}

// ---- dates ------------------------------------------------------------------

const pad = (n: number, width: number): string => String(n).padStart(width, '0');

function formatDateSection(serial: number, tokens: readonly FormatToken[], info: LocaleInfo): string | null {
  const hasTime = tokens.some(
    (tk) => tk.t === 'elapsed' || tk.t === 'ampm' || (tk.t === 'date' && 'hmsf'.includes(tk.field)),
  );
  const when: CivilDateTime | null = DateConversion.tryToDateTime(serial, !hasTime);
  if (!when) return null;
  const twelveHour = tokens.some((tk) => tk.t === 'ampm');
  const totalMs = Math.round(serial * 86_400_000);
  let text = '';
  for (const tk of tokens) {
    switch (tk.t) {
      case 'literal':
        text += tk.text;
        break;
      case 'point':
        text += '.';
        break;
      case 'comma':
        text += ',';
        break;
      case 'percent':
        text += '%';
        break;
      case 'digit':
        text += tk.ch === '?' ? ' ' : tk.ch === '0' ? '0' : '';
        break;
      case 'ampm': {
        const designator = when.hour < 12 ? info.am : info.pm;
        const shown = tk.long ? designator : designator.slice(0, 1);
        text += tk.lower ? shown.toLocaleLowerCase(info.locale) : shown;
        break;
      }
      case 'elapsed': {
        const unitMs = tk.unit === 'h' ? 3_600_000 : tk.unit === 'm' ? 60_000 : 1000;
        text += pad(Math.floor(totalMs / unitMs), tk.length);
        break;
      }
      case 'date':
        text += dateField(tk.field, tk.length, when, twelveHour, info);
        break;
      default:
        break;
    }
  }
  return text;
}

function dateField(
  field: DateField,
  n: number,
  d: CivilDateTime,
  twelveHour: boolean,
  info: LocaleInfo,
): string {
  switch (field) {
    case 'y':
      return n >= 3 ? pad(d.year, 4) : pad(d.year % 100, 2);
    case 'M':
      if (n >= 5) return info.months[d.month - 1]!.slice(0, 1);
      if (n === 4) return info.months[d.month - 1]!;
      if (n === 3) return info.monthsShort[d.month - 1]!;
      return pad(d.month, n);
    case 'd': {
      if (n >= 3) {
        const dow = DateConversion.dayOfWeek(DateConversion.dateToSerial(d.year, d.month, d.day));
        return n >= 4 ? info.days[dow]! : info.daysShort[dow]!;
      }
      return pad(d.day, n);
    }
    case 'h': {
      const hour = twelveHour ? d.hour % 12 || 12 : d.hour;
      return pad(hour, Math.min(n, 2));
    }
    case 'm':
      return pad(d.minute, Math.min(n, 2));
    case 's':
      return pad(d.second, Math.min(n, 2));
    case 'f':
      return pad(d.millisecond, 3).slice(0, n).padEnd(n, '0');
  }
}

// ---- public -----------------------------------------------------------------

const sectionCache = new Map<string, Section[]>();

function sectionsOf(code: string): Section[] {
  let sections = sectionCache.get(code);
  if (!sections) {
    sections = splitSections(code).map(tokenize);
    if (sectionCache.size > 500) sectionCache.clear();
    sectionCache.set(code, sections);
  }
  return sections;
}

/**
 * Renders a `CellValue` for display under an Excel/Calc number-format code and locale. The single
 * formatting path shared by the grid, the status bar, and TEXT(), so they never diverge.
 *
 * Supports sections (`;`, with `[<0]`-style conditions), placeholders `0 # ?`, `,` grouping and
 * scaling, `.`, `%`, `E+00`, quoted/escaped literals, `[$€-407]` currency tokens, colours (ignored),
 * `@`, dates and times (`y m d h s AM/PM A/P`), fractional seconds, and elapsed `[h] [mm] [ss]`.
 * Separators and month/day names come from `Intl` for the locale.
 */
export const NumberFormatter = {
  format(value: CellValue, code: string, locale: string): string {
    switch (value.kind) {
      case 'empty':
        return '';
      case 'error':
        return value.error.token;
      case 'boolean':
        return value.value ? 'TRUE' : 'FALSE';
      case 'text':
        return value.value;
      case 'number':
        break;
    }
    const n = value.value;
    if (!code || code.toLowerCase() === 'general' || code === '@' || !Number.isFinite(n)) {
      return NumberFormatter.formatGeneral(n, locale);
    }
    const info = localeInfo(locale);
    const { section, signed } = pickSection(sectionsOf(code), n);
    if (section.isDate) {
      return formatDateSection(n, section.tokens, info) ?? NumberFormatter.formatGeneral(n, locale);
    }
    const { text, nonZero } = formatNumberSection(Math.abs(n), section.tokens, info);
    return signed && n < 0 && nonZero ? '-' + text : text;
  },

  /**
   * General display: integers plainly (below 1e15), otherwise up to 10 trimmed decimals; a number
   * too small to show at 10 decimals in scientific notation (`1E-12`).
   */
  formatGeneral(n: number, locale: string): string {
    if (Number.isNaN(n)) return 'NaN';
    if (!Number.isFinite(n)) return n > 0 ? 'Infinity' : '-Infinity';
    const sign = n < 0 ? '-' : '';
    if (n === Math.floor(n) && Math.abs(n) < 1e15) return n === 0 ? '0' : sign + String(Math.abs(n));
    // Too small for 10 decimals to show a digit: scientific, as Calc does, rather than a false 0.
    if (Math.abs(n) < GENERAL_SCIENTIFIC_BELOW) return sign + generalScientific(Math.abs(n), locale);
    const { integer, fraction } = fixedDigits(Math.abs(n), 10);
    const trimmed = fraction.replace(/0+$/, '');
    const body = integer + (trimmed ? localeInfo(locale).decimal + trimmed : '');
    return /[1-9]/.test(body) ? sign + body : body;
  },
} as const;

/** Below this, 10 decimals would round a non-zero General number to 0. */
const GENERAL_SCIENTIFIC_BELOW = 5e-11;

/** `1.25E-11`: up to 10 significant digits, trailing zeros trimmed, a two-digit exponent. */
function generalScientific(abs: number, locale: string): string {
  const [mantissa = '0', exponent = '0'] = abs.toExponential(9).split('e');
  const trimmed = mantissa.replace(/\.?0+$/, '').replace('.', localeInfo(locale).decimal);
  const e = Number(exponent);
  return `${trimmed}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`;
}
