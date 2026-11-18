import { Defaults } from '../../defaults.ts';
import { shortDateCode } from '../../editing/inputParser.ts';
import { Argb } from '../../styling/color.ts';
import {
  Alignments,
  BorderEdges,
  Borders,
  CellStyles,
  Fills,
  Fonts,
  NumberFormats,
  type Alignment,
  type Border,
  type BorderEdge,
  type BorderLineStyle,
  type CellStyle,
  type Fill,
  type Font,
  type HorizontalAlignment,
  type VerticalAlignment,
} from '../../styling/cellStyle.ts';
import type { XmlElement } from '../../xml/index.ts';
import { readColor, type XlsxPalette } from './xlsxColors.ts';
import { at, flag, kid, kids } from './xlsxPackage.ts';

/**
 * The number formats Excel numbers 0–49 without writing them into the file (ECMA-376 §18.8.30).
 * 14 and 22 follow the reader's locale, as Excel shows them.
 */
function builtInFormat(id: number, locale: string): string | null {
  switch (id) {
    case 0:
      return 'General';
    case 1:
      return '0';
    case 2:
      return '0.00';
    case 3:
      return '#,##0';
    case 4:
      return '#,##0.00';
    case 9:
      return '0%';
    case 10:
      return '0.00%';
    case 11:
      return '0.00E+00';
    case 12:
      return '# ?/?';
    case 13:
      return '# ??/??';
    case 14:
      return shortDateCode(locale);
    case 15:
      return 'd-mmm-yy';
    case 16:
      return 'd-mmm';
    case 17:
      return 'mmm-yy';
    case 18:
      return 'h:mm AM/PM';
    case 19:
      return 'h:mm:ss AM/PM';
    case 20:
      return 'h:mm';
    case 21:
      return 'h:mm:ss';
    case 22:
      return `${shortDateCode(locale)} h:mm`;
    case 37:
      return '#,##0 ;(#,##0)';
    case 38:
      return '#,##0 ;[Red](#,##0)';
    case 39:
      return '#,##0.00;(#,##0.00)';
    case 40:
      return '#,##0.00;[Red](#,##0.00)';
    case 41:
      return '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)';
    case 42:
      return '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)';
    case 43:
      return '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
    case 44:
      return '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';
    case 45:
      return 'mm:ss';
    case 46:
      return '[h]:mm:ss';
    case 47:
      return 'mm:ss.0';
    case 48:
      return '##0.0E+0';
    case 49:
      return '@';
    default:
      return null;
  }
}

const BORDER_STYLES: Readonly<Record<string, BorderLineStyle>> = {
  thin: 'thin',
  medium: 'medium',
  thick: 'thick',
  dashed: 'dashed',
  dotted: 'dotted',
  double: 'double',
  hair: 'dotted',
  mediumDashed: 'dashed',
  dashDot: 'dashed',
  mediumDashDot: 'dashed',
  dashDotDot: 'dashed',
  mediumDashDotDot: 'dashed',
  slantDashDot: 'dashed',
};

const HORIZONTAL: Readonly<Record<string, HorizontalAlignment>> = {
  general: 'general',
  left: 'left',
  center: 'center',
  right: 'right',
  fill: 'left',
  justify: 'left',
  centerContinuous: 'center',
  distributed: 'left',
};

const VERTICAL: Readonly<Record<string, VerticalAlignment>> = {
  top: 'top',
  center: 'middle',
  bottom: 'bottom',
  justify: 'top',
  distributed: 'middle',
};

/** The cell formats of `styles.xml` (`cellXfs`), as Cellestra styles, by xf index. */
export interface XlsxStyles {
  readonly cellStyles: readonly CellStyle[];
  /** Cell formats whose numbers are really TRUE/FALSE (LibreOffice saves a boolean that way). */
  readonly booleanFormats: ReadonlySet<number>;
}

/** The number format LibreOffice gives a boolean when it saves `.xlsx`: nonzero is TRUE. */
const BOOLEAN_FORMAT = /^"TRUE";"TRUE";"FALSE"$/i;

/** An OOXML boolean child (`<b/>`, `<i val="0"/>`): present means true unless `val` says otherwise. */
function onOff(el: XmlElement | null): boolean {
  return el !== null && flag(at(el, 'val'), true);
}

/**
 * Read `styles.xml` into one Cellestra style per cell format. The workbook's default font (the
 * first, which the Normal style uses) becomes Cellestra's default font, so ordinary cells carry no
 * font style (PRD §8.19); any other font keeps its family and size.
 */
export function readStyles(styles: XmlElement | null, palette: XlsxPalette, locale: string): XlsxStyles {
  const formats = new Map<number, string>();
  for (const f of kids(kid(styles, 'numFmts'), 'numFmt')) {
    const id = Number(at(f, 'numFmtId'));
    const code = at(f, 'formatCode');
    if (Number.isInteger(id) && code !== null) formats.set(id, code);
  }
  const fontEls = kids(kid(styles, 'fonts'), 'font');
  const defaultFont = fontEls[0] ?? null;
  const fonts = fontEls.map((f) => readFont(f, defaultFont, palette));
  const fills = kids(kid(styles, 'fills'), 'fill').map((f) => readFill(f, palette));
  const borders = kids(kid(styles, 'borders'), 'border').map((b) => readBorder(b, palette));

  const booleanFormats = new Set<number>();
  const cellStyles = kids(kid(styles, 'cellXfs'), 'xf').map((xf, index): CellStyle => {
    const numFmtId = Number(at(xf, 'numFmtId') ?? 0);
    let code = formats.get(numFmtId) ?? builtInFormat(numFmtId, locale) ?? 'General';
    if (BOOLEAN_FORMAT.test(code)) {
      booleanFormats.add(index);
      code = 'General';
    }
    return {
      font: fonts[Number(at(xf, 'fontId') ?? 0)] ?? Fonts.default,
      fill: fills[Number(at(xf, 'fillId') ?? 0)] ?? Fills.none,
      border: borders[Number(at(xf, 'borderId') ?? 0)] ?? Borders.none,
      alignment: readAlignment(kid(xf, 'alignment')),
      numberFormat: NumberFormats.isGeneral({ code }) ? NumberFormats.general : NumberFormats.of(code),
    };
  });
  return { cellStyles: cellStyles.length > 0 ? cellStyles : [CellStyles.default], booleanFormats };
}

function readFont(font: XmlElement, defaultFont: XmlElement | null, palette: XlsxPalette): Font {
  const family = at(kid(font, 'name'), 'val');
  const size = Number(at(kid(font, 'sz'), 'val'));
  const sameAsDefault =
    family === at(kid(defaultFont, 'name'), 'val') && size === Number(at(kid(defaultFont, 'sz'), 'val'));
  const underline = kid(font, 'u');
  return {
    family: sameAsDefault || !family ? Defaults.fontFamily : family,
    sizePoints: sameAsDefault || !(size > 0) ? Defaults.fontSizePoints : size,
    bold: onOff(kid(font, 'b')),
    italic: onOff(kid(font, 'i')),
    underline: underline !== null && at(underline, 'val') !== 'none',
    strikethrough: onOff(kid(font, 'strike')),
    color: readColor(kid(font, 'color'), palette) ?? Argb.black,
  };
}

function readFill(fill: XmlElement, palette: XlsxPalette): Fill {
  const pattern = kid(fill, 'patternFill');
  if (pattern) {
    const type = at(pattern, 'patternType') ?? 'none';
    if (type === 'none' || type === 'gray125') return Fills.none;
    // A solid fill's colour is its foreground; other patterns are drawn solid in that colour.
    const color = readColor(kid(pattern, 'fgColor'), palette) ?? readColor(kid(pattern, 'bgColor'), palette);
    return color === null ? Fills.none : Fills.solid(color);
  }
  const stop = kid(kid(fill, 'gradientFill'), 'stop');
  const color = readColor(kid(stop, 'color'), palette);
  return color === null ? Fills.none : Fills.solid(color);
}

function readEdge(edge: XmlElement | null, palette: XlsxPalette): BorderEdge {
  const style = BORDER_STYLES[at(edge, 'style') ?? ''];
  if (!style) return BorderEdges.none;
  return BorderEdges.of(style, readColor(kid(edge, 'color'), palette) ?? Argb.black);
}

function readBorder(border: XmlElement, palette: XlsxPalette): Border {
  const edges = {
    left: readEdge(kid(border, 'left') ?? kid(border, 'start'), palette),
    right: readEdge(kid(border, 'right') ?? kid(border, 'end'), palette),
    top: readEdge(kid(border, 'top'), palette),
    bottom: readEdge(kid(border, 'bottom'), palette),
  };
  return Borders.hasAnyEdge(edges) ? edges : Borders.none;
}

function readAlignment(el: XmlElement | null): Alignment {
  if (!el) return Alignments.default;
  const degrees = Number(at(el, 'textRotation') ?? 0);
  // 1–90 turn counter-clockwise, 91–180 clockwise by (n − 90); 255 is stacked text, not a turn.
  const rotation =
    degrees > 0 && degrees <= 90 ? degrees : degrees > 90 && degrees <= 180 ? 450 - degrees : 0;
  const alignment: Alignment = {
    horizontal: HORIZONTAL[at(el, 'horizontal') ?? ''] ?? 'general',
    vertical: VERTICAL[at(el, 'vertical') ?? ''] ?? 'bottom',
    wrapText: flag(at(el, 'wrapText')),
  };
  return rotation === 0 ? alignment : { ...alignment, rotation };
}
