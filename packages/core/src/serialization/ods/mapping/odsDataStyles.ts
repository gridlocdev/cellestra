import { NumberFormats, type NumberFormat } from '../../../styling/cellStyle.ts';
import { XmlElement } from '../../../xml/index.ts';
import { NumberNs as NumberX, Style } from '../odsXmlNames.ts';

/**
 * Maps Cellestra number-format codes to ODF data styles (`number:*-style`) and back. Covers the
 * common v1 cases — General, fixed decimals, percent, thousands, currency, scientific, fractions,
 * dates; anything else falls back to General so files still open cleanly.
 */
export const OdsDataStyles = {
  /** Build the ODF data-style element for a format, or null for General. */
  build(name: string, format: NumberFormat): XmlElement | null {
    if (NumberFormats.isGeneral(format)) return null;

    // Text ("@") is a style, not the absence of one: it is what stops a cell parsing its content as
    // a number. Returning null here wrote no data style at all, so the cell reopened as Automatic
    // and its content went back to being interpreted.
    if (format.code === '@') {
      return named(NumberX('text-style'), name).add(new XmlElement(NumberX('text-content')));
    }

    switch (NumberFormats.kind(format)) {
      case 'percent':
        return percentage(name, format.code);
      case 'currency':
        return currency(name, format.code) ?? numberStyle(name, format.code);
      case 'scientific':
        return scientific(name, format.code);
      case 'fraction':
        return fraction(name, format.code);
      case 'number':
        return numberStyle(name, format.code);
      case 'date':
      case 'dateTime':
      case 'time':
        return dateStyle(name, format.code);
      default:
        return null;
    }
  },

  /** Reconstruct a format code from an ODF data-style element (best effort). */
  toCode(dataStyle: XmlElement): string {
    switch (dataStyle.local) {
      case 'percentage-style':
        return percentCode(dataStyle);
      case 'number-style':
        return numberCode(dataStyle);
      case 'currency-style':
        return currencyCode(dataStyle);
      case 'date-style':
        return dateCode(dataStyle);
      case 'text-style':
        return '@';
      default:
        return 'General';
    }
  },
} as const;

function named(elementName: ReturnType<typeof NumberX>, name: string): XmlElement {
  return new XmlElement(elementName).setAttr(Style('name'), name);
}

function numberStyle(name: string, code: string): XmlElement {
  const number = new XmlElement(NumberX('number'))
    .setAttr(NumberX('decimal-places'), decimalPlaces(code))
    .setAttr(NumberX('min-integer-digits'), 1);
  if (code.includes(',')) number.setAttr(NumberX('grouping'), 'true');
  return named(NumberX('number-style'), name).add(number);
}

/**
 * A `number:number-style` holding a `number:scientific-number` — ODF's exponent format. Written as a
 * plain `number:number`, `0.00E+00` reopened as `0.00`: the exponent is carried by the element type,
 * and there is nowhere else to put it.
 */
function scientific(name: string, code: string): XmlElement {
  const e = code.toUpperCase().indexOf('E');
  let exponentDigits = 0;
  for (let i = e + 1; i < code.length; i++) {
    const c = code[i]!;
    if (c === '0' || c === '#') exponentDigits++;
    else if (c !== '+' && c !== '-') break;
  }
  return named(NumberX('number-style'), name).add(
    new XmlElement(NumberX('scientific-number'))
      .setAttr(NumberX('decimal-places'), decimalPlaces(code))
      .setAttr(NumberX('min-integer-digits'), integerDigits(code))
      .setAttr(NumberX('min-exponent-digits'), Math.max(1, exponentDigits)),
  );
}

/**
 * A `number:number-style` holding a `number:fraction` (`# ?/?`). Cellestra doesn't render fractions,
 * but a format it can't draw is still a format the file chose — dropping it to General on save would
 * rewrite someone else's document to say something it didn't.
 */
function fraction(name: string, code: string): XmlElement {
  const slash = code.indexOf('/');
  const numerator = countPlaceholders(code, slash < 0 ? 0 : slash - 1, -1);
  const denominator = countPlaceholders(code, slash + 1, +1);
  return named(NumberX('number-style'), name).add(
    new XmlElement(NumberX('fraction'))
      .setAttr(NumberX('min-integer-digits'), code.includes('#') ? 0 : 1)
      .setAttr(NumberX('min-numerator-digits'), numerator)
      .setAttr(NumberX('min-denominator-digits'), denominator),
  );
}

/** Count the run of `?`/`0`/`#` placeholders from `start` in `step`'s direction. */
function countPlaceholders(code: string, start: number, step: number): number {
  let count = 0;
  for (let i = start; i >= 0 && i < code.length && '?0#'.includes(code[i]!); i += step) count++;
  return Math.max(1, count);
}

/** How many leading digit placeholders a code asks for — the `0`s before its decimal point. */
function integerDigits(code: string): number {
  let count = 0;
  for (const c of code) {
    if (c === '.') break;
    if (c === '0') count++;
  }
  return Math.max(1, count);
}

/**
 * A `number:currency-style` — the only ODF style that carries the currency symbol. Writing a currency
 * format as a plain `number:number-style` drops the symbol, so the file reopened as a
 * thousands-separated number: the format survived, the fact that it was money did not.
 *
 * The code is split at its numeric core (the run of `#0,.`): what precedes and follows it is literal
 * text, and whichever side holds the currency symbol becomes `number:currency-symbol`. Returns null
 * when the code has no numeric core to build around, leaving the caller its fallback.
 */
function currency(name: string, code: string): XmlElement | null {
  const start = code.search(/[#0]/);
  if (start < 0) return null;
  let end = start;
  while (end < code.length && '#0,.'.includes(code[end]!)) end++;

  const core = code.slice(start, end);
  const style = named(NumberX('currency-style'), name);
  addLiteral(style, code.slice(0, start));
  style.add(
    new XmlElement(NumberX('number'))
      .setAttr(NumberX('decimal-places'), decimalPlaces(core))
      .setAttr(NumberX('min-integer-digits'), 1)
      .setAttr(NumberX('grouping'), core.includes(',') ? 'true' : 'false'),
  );
  addLiteral(style, code.slice(end));
  return style;
}

/** Append a literal run either side of the number, splitting the currency symbol out of it. */
function addLiteral(style: XmlElement, text: string): void {
  if (text.length === 0) return;
  let pending = '';
  const flush = (): void => {
    if (pending.length > 0) {
      style.add(new XmlElement(NumberX('text'), pending));
      pending = '';
    }
  };
  for (const c of text) {
    if (isCurrencySymbol(c)) {
      flush();
      style.add(new XmlElement(NumberX('currency-symbol'), c));
    } else pending += c;
  }
  flush();
}

function isCurrencySymbol(c: string): boolean {
  return c === '$' || c === '€' || c === '£' || c === '¥' || c === '¤';
}

function percentage(name: string, code: string): XmlElement {
  const number = new XmlElement(NumberX('number'))
    .setAttr(NumberX('decimal-places'), decimalPlaces(code))
    .setAttr(NumberX('min-integer-digits'), 1);
  if (code.includes(',')) number.setAttr(NumberX('grouping'), 'true');
  return named(NumberX('percentage-style'), name).add(number, new XmlElement(NumberX('text'), '%'));
}

function dateStyle(name: string, code: string): XmlElement {
  const style = named(NumberX('date-style'), name);
  // Walk the code, emitting year/month/day/separator pieces for the common tokens.
  const lower = code.toLowerCase();
  let i = 0;
  while (i < lower.length) {
    if (lower.startsWith('yyyy', i)) {
      style.add(year(true));
      i += 4;
    } else if (lower.startsWith('yy', i)) {
      style.add(year(false));
      i += 2;
    } else if (lower.startsWith('mmmm', i)) {
      style.add(month('long', true));
      i += 4;
    } else if (lower.startsWith('mmm', i)) {
      style.add(month('short', true));
      i += 3;
    } else if (lower.startsWith('mm', i)) {
      style.add(month('long', false));
      i += 2;
    } else if (lower.startsWith('m', i)) {
      style.add(month('short', false));
      i += 1;
    } else if (lower.startsWith('dd', i)) {
      style.add(day(true));
      i += 2;
    } else if (lower.startsWith('d', i)) {
      style.add(day(false));
      i += 1;
    } else {
      style.add(new XmlElement(NumberX('text'), code[i]!));
      i += 1;
    }
  }
  return style;
}

const year = (full: boolean): XmlElement =>
  new XmlElement(NumberX('year')).setAttr(NumberX('style'), full ? 'long' : 'short');

const month = (style: string, textual: boolean): XmlElement =>
  new XmlElement(NumberX('month')).setAttr(NumberX('style'), style).setAttr(NumberX('textual'), textual);

const day = (full: boolean): XmlElement =>
  new XmlElement(NumberX('day')).setAttr(NumberX('style'), full ? 'long' : 'short');

function decimalPlaces(code: string): number {
  const dot = code.indexOf('.');
  if (dot < 0) return 0;
  let count = 0;
  for (let i = dot + 1; i < code.length && (code[i] === '0' || code[i] === '#'); i++) count++;
  return count;
}

function numberCode(style: XmlElement): string {
  const frac = style.element(NumberX('fraction'));
  if (frac) {
    const whole = readDigits(frac, 'min-integer-digits', 0) > 0 ? '0' : '#';
    return (
      whole +
      ' ' +
      '?'.repeat(readDigits(frac, 'min-numerator-digits', 1)) +
      '/' +
      '?'.repeat(readDigits(frac, 'min-denominator-digits', 1))
    );
  }

  const sci = style.element(NumberX('scientific-number'));
  if (sci) {
    const decimals = readDecimals(sci);
    let mantissa = '0'.repeat(readDigits(sci, 'min-integer-digits', 1));
    if (decimals > 0) mantissa += '.' + '0'.repeat(decimals);
    return mantissa + 'E+' + '0'.repeat(readDigits(sci, 'min-exponent-digits', 2));
  }

  return numberCore(style.element(NumberX('number')));
}

function readDigits(element: XmlElement, attribute: string, fallback: number): number {
  return parseInteger(element.attr(NumberX(attribute))) ?? fallback;
}

/** The numeric part of a code — `#,##0.00` — from a `number:number` element. */
function numberCore(number: XmlElement | null): string {
  const decimals = readDecimals(number);
  const grouping = number?.attr(NumberX('grouping')) === 'true';
  const integer = grouping ? '#,##0' : '0';
  return decimals > 0 ? integer + '.' + '0'.repeat(decimals) : integer;
}

/**
 * Rebuild a currency code by walking the style's children in document order, so the symbol lands
 * back on the side the file put it. Anything else ODF allows in there (a colour, a negative-value
 * `style:map`) isn't part of the code and is skipped.
 */
function currencyCode(style: XmlElement): string {
  let sb = '';
  for (const child of style.elements()) {
    switch (child.local) {
      case 'currency-symbol':
      case 'text':
        sb += child.value;
        break;
      case 'number':
        sb += numberCore(child);
        break;
    }
  }
  // A currency style with no symbol is just a number; keep the code honest rather than classifying
  // it as money on the strength of the element name alone.
  return sb.length > 0 ? sb : numberCode(style);
}

function percentCode(style: XmlElement): string {
  return numberCore(style.element(NumberX('number'))) + '%';
}

function dateCode(style: XmlElement): string {
  let sb = '';
  for (const child of style.elements()) {
    switch (child.local) {
      case 'year':
        sb += child.attr(NumberX('style')) === 'long' ? 'yyyy' : 'yy';
        break;
      case 'month': {
        // A textual month is a different token, not a longer one: mmmm is "January" where mm is
        // "01". Reading both back as m/mm turned every month name into a number.
        const textual = child.attr(NumberX('textual')) === 'true';
        const longForm = child.attr(NumberX('style')) === 'long';
        sb += textual ? (longForm ? 'mmmm' : 'mmm') : longForm ? 'mm' : 'm';
        break;
      }
      case 'day':
        sb += child.attr(NumberX('style')) === 'long' ? 'dd' : 'd';
        break;
      case 'text':
        sb += child.value;
        break;
    }
  }
  return sb.length > 0 ? sb : 'yyyy-mm-dd';
}

function readDecimals(number: XmlElement | null): number {
  return parseInteger(number?.attr(NumberX('decimal-places'))) ?? 0;
}

/** `int.TryParse(…, NumberStyles.Integer, Invariant)`. */
export function parseInteger(text: string | null | undefined): number | null {
  if (text === null || text === undefined) return null;
  const t = text.trim();
  if (!/^[+-]?\d+$/.test(t)) return null;
  const n = Number(t);
  return Number.isSafeInteger(n) && n <= 2147483647 && n >= -2147483648 ? n : null;
}
