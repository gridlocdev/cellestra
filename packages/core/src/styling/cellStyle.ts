import { Defaults } from '../defaults.ts';
import { Argb, type ArgbColor } from './color.ts';
import { classifyNumberFormat, type NumberFormatKind } from './numberFormatParser.ts';

export type HorizontalAlignment = 'general' | 'left' | 'center' | 'right';
export type VerticalAlignment = 'top' | 'middle' | 'bottom';

/** ODF `style:rotation-align`: the edge of the cell rotated text is anchored to (`none` = Calc's "standard"). */
export type RotationAlign = 'none' | 'bottom' | 'top' | 'center';

export interface Alignment {
  readonly horizontal: HorizontalAlignment;
  readonly vertical: VerticalAlignment;
  readonly wrapText: boolean;
  /** Text rotation in degrees counter-clockwise, in [0, 360). Absent means 0 (not rotated). */
  readonly rotation?: number;
  /** Where rotated text is anchored. Only kept beside a non-zero `rotation`; absent means unstated. */
  readonly rotationAlign?: RotationAlign;
}

/**
 * Cell protection (ODF `style:cell-protect` and `style:print-content`). It only takes effect while
 * the sheet is protected. Absent means `Protections.default`, which is Calc's default.
 */
export interface Protection {
  /** The cell cannot be edited while the sheet is protected. */
  readonly locked: boolean;
  /** The formula is not shown while the sheet is protected. */
  readonly hideFormula: boolean;
  /** Nothing of the cell is shown while the sheet is protected (`hidden-and-protected`). */
  readonly hideAll: boolean;
  /** The cell's content prints. */
  readonly printContent: boolean;
}

/**
 * One `style:map` of the source style (a legacy conditional format): while `condition` holds, the
 * cell shows another style. Kept verbatim from the file; the model does not evaluate it.
 */
export interface CellStyleMap {
  /** `style:condition`, as written in the file (`cell-content()>5`). */
  readonly condition: string;
  /**
   * The common (named) style the condition applies, kept by name because styles.xml is saved as it
   * was. Null when the file named an automatic style, which a save renames: then `applyStyle` is
   * written as a new automatic style and the map points at that.
   */
  readonly applyStyleName: string | null;
  /** What the applied style resolved to on read (without its own maps), or null if it named nothing. */
  readonly applyStyle: CellStyle | null;
  /** `style:base-cell-address`, the cell relative references in `condition` are relative to. */
  readonly baseCellAddress: string | null;
}

export type BorderLineStyle = 'none' | 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double';

export interface BorderEdge {
  readonly style: BorderLineStyle;
  readonly color: ArgbColor;
}

export interface Border {
  readonly left: BorderEdge;
  readonly right: BorderEdge;
  readonly top: BorderEdge;
  readonly bottom: BorderEdge;
}

export interface Font {
  readonly family: string;
  readonly sizePoints: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  readonly strikethrough: boolean;
  readonly color: ArgbColor;
}

/** `background` null = no fill. */
export interface Fill {
  readonly background: ArgbColor | null;
}

/** A number-format code in the Excel/Calc pattern dialect (`#,##0.00`, `0%`, `yyyy-mm-dd`, `@`). */
export interface NumberFormat {
  readonly code: string;
}

/** A cell's complete formatting. Immutable; intern via `StylePool`. */
export interface CellStyle {
  readonly font: Font;
  readonly fill: Fill;
  readonly border: Border;
  readonly alignment: Alignment;
  readonly numberFormat: NumberFormat;
  /**
   * Draw a TRUE/FALSE value as a checkbox that a click toggles. Other values show as usual.
   * Absent means off, so styles made before checkboxes existed need no change.
   */
  readonly checkbox?: boolean;
  /** Absent means `Protections.default`. */
  readonly protection?: Protection;
  /** The source style's `style:map` elements, in order. Absent or empty means none. */
  readonly conditionalMaps?: readonly CellStyleMap[];
}

export const BorderEdges = {
  none: Object.freeze({ style: 'none', color: Argb.black }) as BorderEdge,
  of: (style: BorderLineStyle, color: ArgbColor = Argb.black): BorderEdge => ({ style, color }),
  isVisible: (e: BorderEdge): boolean => e.style !== 'none',
};

export const Borders = {
  none: Object.freeze({
    left: BorderEdges.none,
    right: BorderEdges.none,
    top: BorderEdges.none,
    bottom: BorderEdges.none,
  }) as Border,
  all(style: BorderLineStyle): Border {
    const e = BorderEdges.of(style);
    return { left: e, right: e, top: e, bottom: e };
  },
  hasAnyEdge: (b: Border): boolean =>
    BorderEdges.isVisible(b.left) ||
    BorderEdges.isVisible(b.right) ||
    BorderEdges.isVisible(b.top) ||
    BorderEdges.isVisible(b.bottom),
};

export const Fonts = {
  default: Object.freeze({
    family: Defaults.fontFamily,
    sizePoints: Defaults.fontSizePoints,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    color: Argb.black,
  }) as Font,
};

export const Fills = {
  none: Object.freeze({ background: null }) as Fill,
  solid: (color: ArgbColor): Fill => ({ background: color }),
  hasBackground: (f: Fill): boolean => f.background !== null && !Argb.isTransparent(f.background),
};

export const Alignments = {
  default: Object.freeze({ horizontal: 'general', vertical: 'bottom', wrapText: false }) as Alignment,
};

export const Protections = {
  default: Object.freeze({
    locked: true,
    hideFormula: false,
    hideAll: false,
    printContent: true,
  }) as Protection,
  isDefault: (p: Protection | undefined): boolean =>
    p === undefined || (p.locked && !p.hideFormula && !p.hideAll && p.printContent),
};

export const NumberFormats = {
  general: Object.freeze({ code: 'General' }) as NumberFormat,
  integer: Object.freeze({ code: '0' }) as NumberFormat,
  twoDecimals: Object.freeze({ code: '0.00' }) as NumberFormat,
  percent: Object.freeze({ code: '0%' }) as NumberFormat,
  text: Object.freeze({ code: '@' }) as NumberFormat,
  of: (code: string): NumberFormat => ({ code }),
  isGeneral: (f: NumberFormat): boolean => !f.code || f.code.toLowerCase() === 'general',
  kind: (f: NumberFormat): NumberFormatKind => classifyNumberFormat(f.code),
};

export const CellStyles = {
  default: Object.freeze({
    font: Fonts.default,
    fill: Fills.none,
    border: Borders.none,
    alignment: Alignments.default,
    numberFormat: NumberFormats.general,
  }) as CellStyle,
};

/** A canonical string identity for value-equality interning (field order independent). */
export function styleKey(s: CellStyle): string {
  const f = s.font;
  const e = (x: BorderEdge) => `${x.style}:${x.color >>> 0}`;
  const b = s.border;
  const a = s.alignment;
  const p = s.protection ?? Protections.default;
  return [
    f.family,
    f.sizePoints,
    +f.bold,
    +f.italic,
    +f.underline,
    +f.strikethrough,
    f.color >>> 0,
    s.fill.background === null ? '-' : s.fill.background >>> 0,
    e(b.left),
    e(b.right),
    e(b.top),
    e(b.bottom),
    a.horizontal,
    a.vertical,
    +a.wrapText,
    s.numberFormat.code,
    +!!s.checkbox,
    a.rotation ?? 0,
    a.rotationAlign ?? '',
    `${+p.locked}${+p.hideFormula}${+p.hideAll}${+p.printContent}`,
    mapsKey(s.conditionalMaps),
  ].join('\u0001');
}

function mapsKey(maps: readonly CellStyleMap[] | undefined): string {
  if (!maps || maps.length === 0) return '';
  return JSON.stringify(
    maps.map((m) => [
      m.condition,
      m.applyStyleName,
      m.baseCellAddress,
      m.applyStyle === null ? null : styleKey(m.applyStyle),
    ]),
  );
}
