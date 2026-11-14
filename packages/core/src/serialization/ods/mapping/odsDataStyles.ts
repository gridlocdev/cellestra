import { NumberFormats, type NumberFormat } from '../../../styling/cellStyle.ts';
import type { NumberFormatKind } from '../../../styling/numberFormatParser.ts';
import { splitSections, tokenize, type FormatToken } from '../../../styling/numberFormatter.ts';
import { XmlElement, xmlNamespace } from '../../../xml/index.ts';
import { Fo, NumberNs as NumberX, Style } from '../odsXmlNames.ts';

/**
 * Maps Cellestra number-format codes to ODF data styles (`number:*-style`) and back: General, fixed
 * decimals, percent, thousands, currency (with `[$sym-lcid]` locales), scientific, fractions, dates,
 * times, date-times and elapsed times. A code with sections (`pos;neg`, `pos;neg;zero`, or
 * `[cond]`-guarded ones) becomes one style per section, the last carrying `style:map`s that pick
 * the others — how ODF spells sections. Anything unrecognised falls back to General.
 */
export const OdsDataStyles = {
  /**
   * Build the ODF data-style elements for a format — sub-styles first, the one named `name` last
   * — or an empty list for General.
   */
  build(name: string, format: NumberFormat): XmlElement[] {
    if (NumberFormats.isGeneral(format)) return [];

    // Text ("@") is a style, not the absence of one: it is what stops a cell parsing its content as
    // a number. Returning null here wrote no data style at all, so the cell reopened as Automatic
    // and its content went back to being interpreted.
    if (format.code === '@') {
      return [named(NumberX('text-style'), name).add(new XmlElement(NumberX('text-content')))];
    }

    const kind = NumberFormats.kind(format);
    if (kind === 'date' || kind === 'dateTime' || kind === 'time') {
      // Sections in a date code are rare and ODF date styles cannot be told apart by sign; the
      // first section is the format.
      return [dateTimeStyle(name, splitSections(format.code)[0] ?? format.code, kind)];
    }
    if (
      kind !== 'percent' &&
      kind !== 'currency' &&
      kind !== 'scientific' &&
      kind !== 'fraction' &&
      kind !== 'number'
    ) {
      return [];
    }

    // A fourth section formats text; ODF has no numeric equivalent, so it is dropped.
    const sections = splitSections(format.code).slice(0, 3);
    if (sections.length === 1) return [numericStyle(name, sections[0]!, kind)];

    const parsed = sections.map(parseSection);
    const explicit = parsed.some((p) => p.condition !== null);
    const implied = sections.length === 2 ? ['>=0'] : ['>0', '<0'];
    const out: XmlElement[] = [];
    const main = numericStyle(name, sections.at(-1)!, kind);
    for (let i = 0; i < sections.length - 1; i++) {
      const subName = `${name}P${i}`;
      out.push(numericStyle(subName, sections[i]!, kind).setAttr(Style('volatile'), 'true'));
      const condition = explicit ? (parsed[i]!.condition ?? implied[i] ?? '>=0') : implied[i]!;
      main.add(
        new XmlElement(Style('map'))
          .setAttr(Style('condition'), `value()${condition}`)
          .setAttr(Style('apply-style-name'), subName),
      );
    }
    out.push(main);
    return out;
  },

  /**
   * Reconstruct a format code from an ODF data-style element. `lookup` finds the styles its
   * `style:map`s apply (the other sections); without it only this element's own section is read.
   */
  toCode(dataStyle: XmlElement, lookup: (name: string) => XmlElement | null = () => null): string {
    const own = sectionCode(dataStyle);
    if (own === null) return 'General';

    const maps = dataStyle.elements(Style('map'));
    if (maps.length === 0) return own;

    const subs: Array<{ condition: string; code: string }> = [];
    for (const map of maps) {
      const target = lookup(map.attr(Style('apply-style-name')) ?? '');
      const condition = /^value\(\)\s*(.+)$/.exec((map.attr(Style('condition')) ?? '').trim())?.[1];
      const code = target ? sectionCode(target) : null;
      if (condition === undefined || code === null) continue;
      subs.push({ condition: condition.replaceAll(' ', '').replace('!=', '<>'), code });
    }
    if (subs.length === 0) return own;

    // The two layouts every spreadsheet writes need no brackets: pos;neg and pos;neg;zero.
    const conditions = subs.map((s) => s.condition);
    if (conditions.length === 1 && conditions[0] === '>=0') return `${subs[0]!.code};${own}`;
    if (conditions.length === 2 && conditions[0] === '>0' && conditions[1] === '<0') {
      return `${subs[0]!.code};${subs[1]!.code};${own}`;
    }
    return [...subs.map((s) => `[${s.condition}]${s.code}`), own].join(';');
  },
} as const;

/** LibreOffice's extension namespace; older versions wrote `max-denominator-value` in it. */
const LoExt = xmlNamespace('urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0', 'loext');

function named(elementName: ReturnType<typeof NumberX>, name: string): XmlElement {
  return new XmlElement(elementName).setAttr(Style('name'), name);
}

// ---- write: numeric sections -----------------------------------------------

type Piece = { readonly text: string } | { readonly symbol: string; readonly lcid: string | null };

interface ParsedSection {
  readonly color: string | null;
  readonly condition: string | null;
  readonly prefix: Piece[];
  /** The numeric core, raw: `#,##0.00`, `0.00E+00`, `# ?/?`. */
  readonly core: string;
  readonly suffix: Piece[];
}

const CORE_START = '#0?';
const CORE_CHARS = '#0?,.';
const BARE_SYMBOLS = '$€£¥¤';

/**
 * Split one section into its colour, condition, the literal pieces either side of the number
 * (currency symbols separated out) and the number itself. Quotes, escapes, `_x` and `*x` are
 * literal-text syntax and never reach the core.
 */
function parseSection(section: string): ParsedSection {
  let color: string | null = null;
  let condition: string | null = null;
  const prefix: Piece[] = [];
  const suffix: Piece[] = [];
  let core = '';
  let pendingCore = ''; // literal text inside the core (the space of `# ?/?`, the E+ of 1E+00)
  let inCore = false;
  let coreDone = false;

  const literal = (text: string): void => {
    if (inCore && !coreDone) {
      pendingCore += text;
      return;
    }
    pushText(coreDone ? suffix : prefix, text);
  };
  const symbol = (sym: string, lcid: string | null): void => {
    if (inCore) closeCore();
    (coreDone ? suffix : prefix).push({ symbol: sym, lcid });
  };
  const closeCore = (): void => {
    if (!inCore || coreDone) return;
    coreDone = true;
    if (pendingCore) pushText(suffix, pendingCore);
    pendingCore = '';
  };

  for (let i = 0; i < section.length; i++) {
    const c = section[i]!;
    if (c === '"') {
      const end = section.indexOf('"', i + 1);
      const stop = end < 0 ? section.length : end;
      literal(section.slice(i + 1, stop));
      i = stop;
    } else if (c === '\\') {
      if (i + 1 < section.length) literal(section[i + 1]!);
      i++;
    } else if (c === '_') {
      literal(' ');
      i++;
    } else if (c === '*') {
      i++;
    } else if (c === '[') {
      const end = section.indexOf(']', i + 1);
      const stop = end < 0 ? section.length : end;
      const inside = section.slice(i + 1, stop);
      i = stop;
      if (inside.startsWith('$')) {
        const dash = inside.indexOf('-', 1);
        const sym = inside.slice(1, dash < 0 ? inside.length : dash);
        const lcid = dash < 0 ? null : inside.slice(dash + 1);
        if (sym) symbol(sym, lcid);
      } else if (/^(<=|>=|<>|<|>|=)/.test(inside)) {
        condition = inside.replaceAll(' ', '');
      } else if (COLOR_NAMES[inside.toLowerCase()] !== undefined) {
        color = COLOR_NAMES[inside.toLowerCase()]!;
      }
    } else if (!coreDone && (CORE_START.includes(c) || (inCore && CORE_CHARS.includes(c)))) {
      if (!inCore) inCore = true;
      core += pendingCore + c;
      pendingCore = '';
    } else if (inCore && !coreDone && (c === 'E' || c === 'e') && /[+-]/.test(section[i + 1] ?? '')) {
      core += pendingCore + c + section[i + 1]!;
      pendingCore = '';
      i++;
    } else if (inCore && !coreDone && c === '/') {
      core += pendingCore + c;
      pendingCore = '';
    } else if (inCore && !coreDone && /[1-9]/.test(c) && /\/[0-9]*$/.test(core) && !pendingCore) {
      core += c; // a fixed denominator: the 4 of `?/4`
    } else if (BARE_SYMBOLS.includes(c)) {
      symbol(c, null);
    } else {
      literal(c);
    }
  }
  closeCore();
  return { color, condition, prefix, core, suffix };
}

function pushText(pieces: Piece[], text: string): void {
  if (!text) return;
  const last = pieces.at(-1);
  if (last && 'text' in last) pieces[pieces.length - 1] = { text: last.text + text };
  else pieces.push({ text });
}

function numericStyle(name: string, section: string, kind: NumberFormatKind): XmlElement {
  const parsed = parseSection(section);
  const element =
    kind === 'currency'
      ? named(NumberX('currency-style'), name)
      : kind === 'percent'
        ? named(NumberX('percentage-style'), name)
        : named(NumberX('number-style'), name);

  if (parsed.color !== null)
    element.add(new XmlElement(Style('text-properties')).setAttr(Fo('color'), parsed.color));
  addPieces(element, parsed.prefix, kind);
  // A section of literal text alone (the "-" a zero section often prints) has no number in it.
  if (parsed.core) element.add(numberElement(parsed.core, kind));
  addPieces(element, parsed.suffix, kind);
  return element;
}

function addPieces(style: XmlElement, pieces: readonly Piece[], kind: NumberFormatKind): void {
  for (const piece of pieces) {
    if ('text' in piece) {
      style.add(new XmlElement(NumberX('text'), piece.text));
    } else if (kind === 'currency') {
      // The only ODF element that carries a currency symbol, with its locale when the code had one.
      const symbol = new XmlElement(NumberX('currency-symbol'), piece.symbol);
      const locale = piece.lcid === null ? null : LCID_TO_LOCALE[piece.lcid.toUpperCase().replace(/^0+/, '')];
      if (locale) {
        symbol.setAttr(NumberX('language'), locale[0]);
        symbol.setAttr(NumberX('country'), locale[1]);
      }
      style.add(symbol);
    } else {
      style.add(new XmlElement(NumberX('text'), piece.symbol));
    }
  }
}

function numberElement(core: string, kind: NumberFormatKind): XmlElement {
  if (kind === 'scientific' && /e/i.test(core)) {
    const e = core.toUpperCase().indexOf('E');
    let exponentDigits = 0;
    for (let i = e + 1; i < core.length; i++) {
      const c = core[i]!;
      if (c === '0' || c === '#') exponentDigits++;
      else if (c !== '+' && c !== '-') break;
    }
    // ODF's exponent format: the exponent is carried by the element type, nowhere else.
    return new XmlElement(NumberX('scientific-number'))
      .setAttr(NumberX('decimal-places'), decimalPlaces(core.slice(0, e)))
      .setAttr(NumberX('min-integer-digits'), integerDigits(core))
      .setAttr(NumberX('min-exponent-digits'), Math.max(1, exponentDigits));
  }
  if (kind === 'fraction' && core.includes('/')) return fractionElement(core);
  const decimals = decimalPlaces(core);
  const number = new XmlElement(NumberX('number'))
    .setAttr(NumberX('decimal-places'), decimals)
    .setAttr(NumberX('min-integer-digits'), integerZeros(core));
  // `0.0#`: one decimal always shown, a second only when it is not zero.
  const minDecimals = minDecimalPlaces(core);
  if (minDecimals !== decimals) number.setAttr(NumberX('min-decimal-places'), minDecimals);
  if (core.includes(',') || kind === 'currency')
    number.setAttr(NumberX('grouping'), core.includes(',') ? 'true' : 'false');
  return number;
}

/**
 * `# ?/?`, `0 ??/??`, `?/4` as an ODF `number:fraction`. A whole-number part is its
 * `min-integer-digits` (0 for `#`); leaving the attribute out means there is none, as LibreOffice
 * reads it. A fixed denominator is `denominator-value`; otherwise `max-denominator-value` says how
 * many digits the denominator may have.
 */
function fractionElement(core: string): XmlElement {
  const slash = core.indexOf('/');
  const before = core.slice(0, slash).trimEnd();
  const numerator = /[?0#]+$/.exec(before)?.[0] ?? '?';
  const integer = before.slice(0, before.length - numerator.length).trim();
  const after = core.slice(slash + 1).trimStart();
  const fixed = /^[1-9][0-9]*/.exec(after)?.[0];
  const denominatorDigits = fixed ? fixed.length : (/^[?0#]+/.exec(after)?.[0].length ?? 1);
  const element = new XmlElement(NumberX('fraction'));
  if (integer) element.setAttr(NumberX('min-integer-digits'), [...integer].filter((c) => c === '0').length);
  element
    .setAttr(NumberX('min-numerator-digits'), numerator.length)
    .setAttr(NumberX('min-denominator-digits'), denominatorDigits);
  if (fixed) element.setAttr(NumberX('denominator-value'), fixed);
  else if (denominatorDigits > 1)
    element.setAttr(NumberX('max-denominator-value'), 10 ** denominatorDigits - 1);
  return element;
}

/** How many leading digit placeholders a code asks for — the `0`s before its decimal point. */
function integerDigits(code: string): number {
  let count = 0;
  for (const c of code) {
    if (c === '.' || c === 'E' || c === 'e') break;
    if (c === '0') count++;
  }
  return Math.max(1, count);
}

/** The `0`s before the decimal point of a plain number core: `00000` → 5, `#,###` → 0. */
function integerZeros(code: string): number {
  let count = 0;
  for (const c of code) {
    if (c === '.') break;
    if (c === '0') count++;
  }
  return count;
}

/** The `0`s after the decimal point: the decimals always shown (`0.0#` → 1). */
function minDecimalPlaces(code: string): number {
  const dot = code.indexOf('.');
  if (dot < 0) return 0;
  let count = 0;
  for (let i = dot + 1; i < code.length && (code[i] === '0' || code[i] === '#' || code[i] === '?'); i++)
    if (code[i] === '0') count++;
  return count;
}

function decimalPlaces(code: string): number {
  const dot = code.indexOf('.');
  if (dot < 0) return 0;
  let count = 0;
  for (let i = dot + 1; i < code.length && (code[i] === '0' || code[i] === '#' || code[i] === '?'); i++)
    count++;
  return count;
}

// ---- write: dates and times -------------------------------------------------

/**
 * A `number:date-style` (dates and date-times) or `number:time-style` (times), one ODF element per
 * token. Writing hours, minutes and seconds through the date tokens turned `hh:mm:ss` into hours,
 * a *month* and seconds; each time field has its own element. An elapsed `[h]` is the style's
 * `number:truncate-on-overflow="false"`.
 */
function dateTimeStyle(name: string, section: string, kind: NumberFormatKind): XmlElement {
  const style = named(NumberX(kind === 'time' ? 'time-style' : 'date-style'), name);
  const { tokens } = tokenize(section);
  let pendingText = '';
  const flush = (): void => {
    if (pendingText) style.add(new XmlElement(NumberX('text'), pendingText));
    pendingText = '';
  };
  const add = (element: XmlElement): void => {
    flush();
    style.add(element);
  };
  const long = (field: XmlElement, isLong: boolean): XmlElement =>
    field.setAttr(NumberX('style'), isLong ? 'long' : 'short');

  for (let i = 0; i < tokens.length; i++) {
    const tk: FormatToken = tokens[i]!;
    switch (tk.t) {
      case 'date':
        switch (tk.field) {
          case 'y':
            add(long(new XmlElement(NumberX('year')), tk.length >= 3));
            break;
          case 'M':
            add(
              long(new XmlElement(NumberX('month')), tk.length === 2 || tk.length >= 4).setAttr(
                NumberX('textual'),
                tk.length >= 3 ? 'true' : 'false',
              ),
            );
            break;
          case 'd':
            add(
              tk.length >= 3
                ? long(new XmlElement(NumberX('day-of-week')), tk.length >= 4)
                : long(new XmlElement(NumberX('day')), tk.length === 2),
            );
            break;
          case 'h':
            add(long(new XmlElement(NumberX('hours')), tk.length >= 2));
            break;
          case 'm':
            add(long(new XmlElement(NumberX('minutes')), tk.length >= 2));
            break;
          case 's': {
            const seconds = long(new XmlElement(NumberX('seconds')), tk.length >= 2);
            const next = tokens[i + 1];
            if (next && next.t === 'date' && next.field === 'f') {
              seconds.setAttr(NumberX('decimal-places'), next.length);
              i++;
            }
            add(seconds);
            break;
          }
          case 'f':
            break;
        }
        break;
      case 'elapsed': {
        style.setAttr(NumberX('truncate-on-overflow'), 'false');
        const field = tk.unit === 'h' ? 'hours' : tk.unit === 'm' ? 'minutes' : 'seconds';
        const element = long(new XmlElement(NumberX(field)), tk.length >= 2);
        const next = tokens[i + 1];
        if (tk.unit === 's' && next && next.t === 'date' && next.field === 'f') {
          element.setAttr(NumberX('decimal-places'), next.length);
          i++;
        }
        add(element);
        break;
      }
      case 'ampm':
        add(new XmlElement(NumberX('am-pm')));
        break;
      case 'literal':
        pendingText += tk.text;
        break;
      case 'point':
        pendingText += '.';
        break;
      case 'comma':
        pendingText += ',';
        break;
      case 'percent':
        pendingText += '%';
        break;
      case 'digit':
        pendingText += tk.ch;
        break;
      default:
        break;
    }
  }
  flush();
  return style;
}

// ---- read -------------------------------------------------------------------

/** The code of one section (one data-style element), or null when it is not a format we read. */
function sectionCode(style: XmlElement): string | null {
  const color = colorCode(style);
  let body: string | null;
  switch (style.local) {
    case 'number-style':
    case 'currency-style':
    case 'percentage-style':
      body = numericCode(style);
      break;
    case 'date-style':
    case 'time-style':
      body = dateTimeCode(style);
      break;
    case 'text-style':
      body = '@';
      break;
    default:
      body = null;
  }
  return body === null ? null : color + body;
}

function colorCode(style: XmlElement): string {
  const hex = style.element(Style('text-properties'))?.attr(Fo('color'))?.toLowerCase();
  if (!hex) return '';
  const name = Object.entries(COLOR_NAMES).find(([, v]) => v === hex)?.[0];
  return name ? `[${name.toUpperCase()}]` : '';
}

/**
 * Walk a numeric style's children in document order — literal text, currency symbols and the
 * number — so each lands back where the file put it.
 */
function numericCode(style: XmlElement): string {
  let code = '';
  let sawNumber = false;
  const percent = style.local === 'percentage-style';
  for (const child of style.elements()) {
    if (child.uri !== NumberX.uri) continue;
    switch (child.local) {
      case 'text':
        code += quoteLiteral(child.value, percent ? '%' : '');
        break;
      case 'currency-symbol':
        code += currencySymbolCode(child);
        break;
      case 'number':
        code += numberCore(child);
        sawNumber = true;
        break;
      case 'scientific-number': {
        const decimals = readDigits(child, 'decimal-places', 0);
        let mantissa = '0'.repeat(readDigits(child, 'min-integer-digits', 1));
        if (decimals > 0) mantissa += '.' + '0'.repeat(decimals);
        code += mantissa + 'E+' + '0'.repeat(readDigits(child, 'min-exponent-digits', 2));
        sawNumber = true;
        break;
      }
      case 'fraction':
        code += fractionCode(child);
        sawNumber = true;
        break;
      case 'text-content':
        code += '@';
        break;
    }
  }
  if (!sawNumber && code.length === 0) code = percent ? '0%' : '0';
  // A currency style with no symbol is just a number, which its code says honestly.
  return code;
}

/** `# ?/?`, `?/4`, `# ??/??` from a `number:fraction` element (see `fractionElement`). */
function fractionCode(fraction: XmlElement): string {
  const hasInteger = fraction.attr(NumberX('min-integer-digits')) !== null;
  const integerZeros = Math.min(readDigits(fraction, 'min-integer-digits', 0), MAX_DIGITS);
  const whole = !hasInteger ? '' : (integerZeros === 0 ? '#' : '0'.repeat(integerZeros)) + ' ';
  const numerator = '?'.repeat(
    Math.max(1, Math.min(readDigits(fraction, 'min-numerator-digits', 1), MAX_DIGITS)),
  );
  const fixed = readDigits(fraction, 'denominator-value', 0);
  if (fixed > 0) return `${whole}${numerator}/${fixed}`;
  // LibreOffice spells the denominator's width as its largest value (99 for ??).
  const maxValue =
    fraction.attr(NumberX('max-denominator-value')) ?? fraction.attr(LoExt('max-denominator-value'));
  const fromMax = maxValue && /^[1-9][0-9]*$/.test(maxValue) ? maxValue.length : 0;
  const width = Math.max(
    1,
    Math.min(Math.max(readDigits(fraction, 'min-denominator-digits', 1), fromMax), MAX_DIGITS),
  );
  return `${whole}${numerator}/${'?'.repeat(width)}`;
}

/** The numeric part of a code — `#,##0.00` — from a `number:number` element. */
function numberCore(number: XmlElement): string {
  const decimals = Math.min(readDigits(number, 'decimal-places', 0), MAX_DIGITS);
  const minDecimals = Math.min(readDigits(number, 'min-decimal-places', decimals), decimals);
  const grouping = number.attr(NumberX('grouping')) === 'true';
  const minInteger = Math.min(readDigits(number, 'min-integer-digits', 1), MAX_DIGITS);
  let integer: string;
  if (grouping) {
    // `#,##0`, `#,###`, `00,000`: the zeros padded with `#` to one full group, a comma before the last three.
    const digits = '#'.repeat(Math.max(0, 4 - minInteger)) + '0'.repeat(minInteger);
    integer = digits.slice(0, -3) + ',' + digits.slice(-3);
  } else {
    integer = minInteger === 0 ? '#' : '0'.repeat(minInteger);
  }
  if (decimals === 0) return integer;
  return integer + '.' + '0'.repeat(minDecimals) + '#'.repeat(decimals - minDecimals);
}

/** A file asking for more digits than any number has is capped, not expanded into a huge code. */
const MAX_DIGITS = 30;

function currencySymbolCode(symbol: XmlElement): string {
  const text = symbol.value;
  const language = symbol.attr(NumberX('language'));
  const country = symbol.attr(NumberX('country'));
  const lcid =
    language && country ? LOCALE_TO_LCID[`${language.toLowerCase()}-${country.toUpperCase()}`] : undefined;
  if (lcid) return `[$${text}-${lcid}]`;
  if (text.length === 1 && BARE_SYMBOLS.includes(text)) return text;
  return `[$${text}]`;
}

function dateTimeCode(style: XmlElement): string {
  const elapsed = style.attr(NumberX('truncate-on-overflow')) === 'false';
  let firstTime = true;
  let code = '';
  for (const child of style.elements()) {
    if (child.uri !== NumberX.uri) continue;
    const isLong = child.attr(NumberX('style')) === 'long';
    switch (child.local) {
      case 'year':
        code += isLong ? 'yyyy' : 'yy';
        break;
      case 'month': {
        // A textual month is a different token, not a longer one: mmmm is "January" where mm is
        // "01". Reading both back as m/mm turned every month name into a number.
        const textual = child.attr(NumberX('textual')) === 'true';
        code += textual ? (isLong ? 'mmmm' : 'mmm') : isLong ? 'mm' : 'm';
        break;
      }
      case 'day':
        code += isLong ? 'dd' : 'd';
        break;
      case 'day-of-week':
        code += isLong ? 'dddd' : 'ddd';
        break;
      case 'hours':
      case 'minutes':
      case 'seconds': {
        let token = child.local[0]!.repeat(isLong ? 2 : 1);
        // An elapsed style lets its leading field run past 24h/60m: [h]:mm:ss.
        if (elapsed && firstTime) token = `[${token}]`;
        firstTime = false;
        code += token;
        if (child.local === 'seconds') {
          const decimals = readDigits(child, 'decimal-places', 0);
          if (decimals > 0) code += '.' + '0'.repeat(decimals);
        }
        break;
      }
      case 'am-pm':
        code += 'AM/PM';
        break;
      case 'text':
        // In a date a comma or full stop is only ever a separator.
        code += quoteLiteral(child.value, ',.');
        break;
    }
  }
  if (code.length > 0) return code;
  return style.local === 'time-style' ? 'hh:mm:ss' : 'yyyy-mm-dd';
}

/**
 * Literal text as format-code syntax: separators stay bare, anything a format would read as a
 * token (letters, `0`, `#`, `.`, `$`…) goes in quotes, with a quote itself escaped.
 */
function quoteLiteral(text: string, alsoBare: string): string {
  let out = '';
  let quoted = '';
  const flush = (): void => {
    if (quoted) out += `"${quoted}"`;
    quoted = '';
  };
  for (const c of text) {
    if (' -/:()'.includes(c) || alsoBare.includes(c)) {
      flush();
      out += c;
    } else if (c === '"') {
      flush();
      out += '\\"';
    } else quoted += c;
  }
  flush();
  return out;
}

function readDigits(element: XmlElement, attribute: string, fallback: number): number {
  return parseInteger(element.attr(NumberX(attribute))) ?? fallback;
}

/** `int.TryParse(…, NumberStyles.Integer, Invariant)`. */
export function parseInteger(text: string | null | undefined): number | null {
  if (text === null || text === undefined) return null;
  const t = text.trim();
  if (!/^[+-]?\d+$/.test(t)) return null;
  const n = Number(t);
  return Number.isSafeInteger(n) && n <= 2147483647 && n >= -2147483648 ? n : null;
}

// ---- tables -----------------------------------------------------------------

/** The eight colours a format code can name, as ODF writes them. */
const COLOR_NAMES: Record<string, string> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00ffff',
  green: '#00ff00',
  magenta: '#ff00ff',
  red: '#ff0000',
  white: '#ffffff',
  yellow: '#ffff00',
};

/** Windows LCIDs (hex, as `[$€-407]` writes them) of the common currency locales. */
const LCID_TO_LOCALE: Record<string, readonly [language: string, country: string]> = {
  '409': ['en', 'US'],
  '809': ['en', 'GB'],
  C09: ['en', 'AU'],
  '1009': ['en', 'CA'],
  '1409': ['en', 'NZ'],
  '1809': ['en', 'IE'],
  '4009': ['en', 'IN'],
  '1C09': ['en', 'ZA'],
  '407': ['de', 'DE'],
  C07: ['de', 'AT'],
  '807': ['de', 'CH'],
  '40C': ['fr', 'FR'],
  '80C': ['fr', 'BE'],
  C0C: ['fr', 'CA'],
  '100C': ['fr', 'CH'],
  '410': ['it', 'IT'],
  C0A: ['es', 'ES'],
  '40A': ['es', 'ES'],
  '80A': ['es', 'MX'],
  '413': ['nl', 'NL'],
  '813': ['nl', 'BE'],
  '816': ['pt', 'PT'],
  '416': ['pt', 'BR'],
  '41D': ['sv', 'SE'],
  '414': ['nb', 'NO'],
  '406': ['da', 'DK'],
  '40B': ['fi', 'FI'],
  '415': ['pl', 'PL'],
  '405': ['cs', 'CZ'],
  '40E': ['hu', 'HU'],
  '419': ['ru', 'RU'],
  '422': ['uk', 'UA'],
  '41F': ['tr', 'TR'],
  '408': ['el', 'GR'],
  '40D': ['he', 'IL'],
  '411': ['ja', 'JP'],
  '804': ['zh', 'CN'],
  '404': ['zh', 'TW'],
  C04: ['zh', 'HK'],
  '412': ['ko', 'KR'],
  '41E': ['th', 'TH'],
  '439': ['hi', 'IN'],
};

const LOCALE_TO_LCID: Record<string, string> = {};
for (const [lcid, [language, country]] of Object.entries(LCID_TO_LOCALE)) {
  LOCALE_TO_LCID[`${language}-${country}`] ??= lcid;
}
