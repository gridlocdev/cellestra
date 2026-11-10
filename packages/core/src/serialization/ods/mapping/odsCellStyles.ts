import { Defaults } from '../../../defaults.ts';
import {
  BorderEdges,
  Borders,
  CellStyles,
  Fills,
  Protections,
  type Alignment,
  type Border,
  type BorderEdge,
  type BorderLineStyle,
  type CellStyle,
  type CellStyleMap,
  type Fill,
  type Font,
  type HorizontalAlignment,
  type NumberFormat,
  type Protection,
  type RotationAlign,
  type VerticalAlignment,
} from '../../../styling/cellStyle.ts';
import { Argb, type ArgbColor } from '../../../styling/color.ts';
import { XmlElement, type XmlName } from '../../../xml/index.ts';
import { Fo, Style } from '../odsXmlNames.ts';

/** A `style:map` as written: the style name it applies is the one the writer resolved. */
export interface WrittenStyleMap {
  readonly condition: string;
  readonly applyStyleName: string;
  readonly baseCellAddress: string | null;
}

/**
 * How a `style:map`'s `style:apply-style-name` resolves on read: the style it names, and whether
 * that is a common (named) style, which keeps its name through a save.
 */
export interface MapStyleTarget {
  readonly style: CellStyle;
  readonly common: boolean;
}

/**
 * The prefix of the common styles a save writes for conditional formats whose applied style has no
 * name of its own (LibreOffice applies only named styles). They are dropped from styles.xml and
 * written again on every save, so edited or deleted rules leave none behind.
 */
export const CONDITIONAL_STYLE_PREFIX = 'CellestraCond';

/** Maps a `CellStyle` to an ODF table-cell automatic style and back. */
export const OdsCellStyles = {
  /** `maps` are the style's `conditionalMaps` with each applied style's name as it is written. */
  build(
    name: string,
    style: CellStyle,
    dataStyleName: string | null,
    maps: readonly WrittenStyleMap[] = [],
  ): XmlElement {
    const element = new XmlElement(Style('style'));
    element.setAttr(Style('name'), name).setAttr(Style('family'), 'table-cell');
    if (dataStyleName !== null) element.setAttr(Style('data-style-name'), dataStyleName);

    const cellProps = new XmlElement(Style('table-cell-properties'));
    if (Fills.hasBackground(style.fill)) {
      cellProps.setAttr(Fo('background-color'), Argb.toHex(style.fill.background!));
    }
    cellProps.setAttr(Style('vertical-align'), VERTICAL_TO_ODF[style.alignment.vertical]);
    if (style.alignment.wrapText) cellProps.setAttr(Fo('wrap-option'), 'wrap');
    addBorders(cellProps, style.border);
    // ODF schema order: common-rotation-angle, style:rotation-align, style:cell-protect, style:print-content.
    const rotation = style.alignment.rotation ?? 0;
    if (rotation !== 0) {
      cellProps.setAttr(Style('rotation-angle'), formatAngle(rotation));
      if (style.alignment.rotationAlign !== undefined) {
        cellProps.setAttr(Style('rotation-align'), style.alignment.rotationAlign);
      }
    }
    if (!Protections.isDefault(style.protection)) {
      const protection = style.protection!;
      cellProps.setAttr(Style('cell-protect'), cellProtectToOdf(protection));
      if (!protection.printContent) cellProps.setAttr(Style('print-content'), 'false');
    }
    if (cellProps.hasAttributes) element.add(cellProps);

    if (style.alignment.horizontal !== 'general') {
      element.add(
        new XmlElement(Style('paragraph-properties')).setAttr(
          Fo('text-align'),
          HORIZONTAL_TO_ODF[style.alignment.horizontal],
        ),
      );
    }

    const textProps = new XmlElement(Style('text-properties'));
    const font = style.font;
    if (font.bold) textProps.setAttr(Fo('font-weight'), 'bold');
    if (font.italic) textProps.setAttr(Fo('font-style'), 'italic');
    if (font.underline) {
      textProps.setAttr(Style('text-underline-style'), 'solid');
      textProps.setAttr(Style('text-underline-width'), 'auto');
      textProps.setAttr(Style('text-underline-color'), 'font-color');
    }
    if (font.strikethrough) textProps.setAttr(Style('text-line-through-style'), 'solid');
    if (font.color >>> 0 !== Argb.black && !Argb.isTransparent(font.color)) {
      textProps.setAttr(Fo('color'), Argb.toHex(font.color));
    }
    if (font.family !== Defaults.fontFamily) textProps.setAttr(Fo('font-family'), font.family);
    if (font.sizePoints !== Defaults.fontSizePoints)
      textProps.setAttr(Fo('font-size'), `${font.sizePoints}pt`);
    if (textProps.hasAttributes) element.add(textProps);

    // style:map follows the property elements (ODF 16.2).
    for (const map of maps) {
      const mapElement = new XmlElement(Style('map'))
        .setAttr(Style('condition'), map.condition)
        .setAttr(Style('apply-style-name'), map.applyStyleName);
      if (map.baseCellAddress !== null) mapElement.setAttr(Style('base-cell-address'), map.baseCellAddress);
      element.add(mapElement);
    }

    return element;
  },

  /**
   * Read one `style:style`'s properties over `base` (the parent style's result, for a
   * `style:parent-style-name` chain): only what this element states changes. `fontFace` resolves a
   * `style:font-name` through the document's `office:font-face-decls`. `mapStyle` resolves a
   * `style:map`'s `style:apply-style-name` (a name it cannot resolve is kept as it is); without it
   * the maps are not read. A style's maps are its own: they are not inherited from `base`.
   */
  parse(
    styleElement: XmlElement,
    resolveDataStyle: (name: string) => NumberFormat,
    base: CellStyle = CellStyles.default,
    fontFace: (name: string) => string | null = () => null,
    mapStyle: ((name: string) => MapStyleTarget | null) | null = null,
  ): CellStyle {
    let font: Font = base.font;
    let fill: Fill = base.fill;
    let alignment: Alignment = base.alignment;
    let border: Border = base.border;
    let protection: Protection = base.protection ?? Protections.default;

    const textProps = styleElement.element(Style('text-properties'));
    if (textProps) {
      const weight = textProps.attr(Fo('font-weight'));
      if (weight !== null) font = { ...font, bold: weight === 'bold' || Number(weight) >= 600 };
      const fontStyle = textProps.attr(Fo('font-style'));
      if (fontStyle !== null) font = { ...font, italic: fontStyle === 'italic' || fontStyle === 'oblique' };
      const underline = textProps.attr(Style('text-underline-style'));
      if (underline !== null) font = { ...font, underline: underline !== 'none' };
      const strike = textProps.attr(Style('text-line-through-style'));
      if (strike !== null) font = { ...font, strikethrough: strike !== 'none' };
      const color = Argb.tryFromHex(textProps.attr(Fo('color')));
      if (color !== null) font = { ...font, color };
      const fontName = textProps.attr(Style('font-name'));
      const family = textProps.attr(Fo('font-family')) ?? (fontName === null ? null : fontFace(fontName));
      if (family) font = { ...font, family: unquoteFamily(family) };
      const size = parsePt(textProps.attr(Fo('font-size')));
      if (size !== null) font = { ...font, sizePoints: size };
    }

    const cellProps = styleElement.element(Style('table-cell-properties'));
    if (cellProps) {
      const background = cellProps.attr(Fo('background-color'));
      if (background === 'transparent') fill = Fills.none;
      else {
        const bg = Argb.tryFromHex(background);
        if (bg !== null) fill = Fills.solid(bg);
      }

      const verticalAttr = cellProps.attr(Style('vertical-align'));
      if (verticalAttr !== null) {
        alignment = { ...alignment, vertical: VERTICAL_FROM_ODF[verticalAttr] ?? 'bottom' };
      }
      const wrap = cellProps.attr(Fo('wrap-option'));
      if (wrap !== null) alignment = { ...alignment, wrapText: wrap === 'wrap' };
      border = parseBorders(cellProps, border);

      const angle = parseAngle(cellProps.attr(Style('rotation-angle')));
      if (angle !== null) alignment = { ...alignment, rotation: angle };
      const rotationAlign = cellProps.attr(Style('rotation-align'));
      if (rotationAlign !== null && ROTATION_ALIGNS.has(rotationAlign)) {
        alignment = { ...alignment, rotationAlign: rotationAlign as RotationAlign };
      }
      protection = parseCellProtect(cellProps.attr(Style('cell-protect')), protection);
      const print = cellProps.attr(Style('print-content'));
      if (print !== null) protection = { ...protection, printContent: print !== 'false' };
    }

    const align = styleElement.element(Style('paragraph-properties'))?.attr(Fo('text-align')) ?? null;
    if (align !== null) {
      alignment = { ...alignment, horizontal: HORIZONTAL_FROM_ODF[align] ?? 'general' };
    }

    const dataStyleName = styleElement.attr(Style('data-style-name'));
    const numberFormat = dataStyleName === null ? base.numberFormat : resolveDataStyle(dataStyleName);

    const style: CellStyle = { font, fill, border, alignment: normalizeRotation(alignment), numberFormat };
    const maps = (mapStyle === null ? [] : styleElement.elements(Style('map'))).flatMap(
      (map): CellStyleMap[] => {
        const condition = map.attr(Style('condition'));
        const name = map.attr(Style('apply-style-name'));
        if (condition === null || name === null) return [];
        const target = mapStyle!(name);
        return [
          {
            condition,
            // A style an earlier save generated has no name worth keeping: the next save writes it again.
            applyStyleName:
              target === null || (target.common && !name.startsWith(CONDITIONAL_STYLE_PREFIX)) ? name : null,
            applyStyle: target?.style ?? null,
            baseCellAddress: map.attr(Style('base-cell-address')),
          },
        ];
      },
    );
    return {
      ...style,
      ...(Protections.isDefault(protection) ? {} : { protection }),
      ...(maps.length > 0 ? { conditionalMaps: maps } : {}),
    };
  },
} as const;

const ROTATION_ALIGNS: ReadonlySet<string> = new Set<RotationAlign>(['none', 'bottom', 'top', 'center']);

/** A zero rotation is no rotation: its anchor means nothing, so neither is kept. */
function normalizeRotation(alignment: Alignment): Alignment {
  if (alignment.rotation === undefined && alignment.rotationAlign === undefined) return alignment;
  const { rotation, rotationAlign, ...rest } = alignment;
  if (!rotation) return rest;
  return rotationAlign === undefined ? { ...rest, rotation } : { ...rest, rotation, rotationAlign };
}

/** An ODF angle (`90`, `90deg`, `1.5708rad`, `100grad`) in degrees within [0, 360), or null. */
function parseAngle(value: string | null): number | null {
  if (value === null) return null;
  const match = /^\s*([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)\s*(deg|rad|grad)?\s*$/.exec(value);
  if (!match) return null;
  const n = Number(match[1]);
  const degrees = match[2] === 'rad' ? (n * 180) / Math.PI : match[2] === 'grad' ? n * 0.9 : n;
  // Calc keeps hundredths of a degree.
  const normalized = Math.round((((degrees % 360) + 360) % 360) * 100) / 100;
  return normalized === 360 ? 0 : normalized;
}

function formatAngle(degrees: number): string {
  return String(Math.round(degrees * 100) / 100);
}

/**
 * `style:cell-protect`: `none`, `hidden-and-protected`, or a space-separated list of `protected`
 * and `formula-hidden`. An unknown value leaves `base` as it was.
 */
function parseCellProtect(value: string | null, base: Protection): Protection {
  if (value === null) return base;
  const tokens = value.trim().split(/\s+/);
  if (tokens.length === 1 && tokens[0] === 'none') {
    return { ...base, locked: false, hideFormula: false, hideAll: false };
  }
  if (tokens.length === 1 && tokens[0] === 'hidden-and-protected') {
    return { ...base, locked: true, hideFormula: true, hideAll: true };
  }
  if (!tokens.every((t) => t === 'protected' || t === 'formula-hidden')) return base;
  return {
    ...base,
    locked: tokens.includes('protected'),
    hideFormula: tokens.includes('formula-hidden'),
    hideAll: false,
  };
}

function cellProtectToOdf(p: Protection): string {
  if (p.hideAll) return 'hidden-and-protected';
  const tokens = [...(p.locked ? ['protected'] : []), ...(p.hideFormula ? ['formula-hidden'] : [])];
  return tokens.length === 0 ? 'none' : tokens.join(' ');
}

const VERTICAL_TO_ODF: Record<VerticalAlignment, string> = { top: 'top', middle: 'middle', bottom: 'bottom' };
const VERTICAL_FROM_ODF: Record<string, VerticalAlignment> = { top: 'top', middle: 'middle' };
const HORIZONTAL_TO_ODF: Record<HorizontalAlignment, string> = {
  general: 'start',
  left: 'start',
  center: 'center',
  right: 'end',
};
const HORIZONTAL_FROM_ODF: Record<string, HorizontalAlignment> = {
  center: 'center',
  end: 'right',
  start: 'left',
};

function addBorders(cellProps: XmlElement, border: Border): void {
  if (!Borders.hasAnyEdge(border)) return;
  addEdge(cellProps, Fo('border-left'), border.left);
  addEdge(cellProps, Fo('border-right'), border.right);
  addEdge(cellProps, Fo('border-top'), border.top);
  addEdge(cellProps, Fo('border-bottom'), border.bottom);
}

function addEdge(cellProps: XmlElement, attribute: XmlName, edge: BorderEdge): void {
  if (BorderEdges.isVisible(edge)) cellProps.setAttr(attribute, borderToOdf(edge));
}

const BORDER_TO_ODF: Record<BorderLineStyle, [width: string, style: string]> = {
  none: ['0.5pt', 'solid'],
  thin: ['0.5pt', 'solid'],
  medium: ['1pt', 'solid'],
  thick: ['2pt', 'solid'],
  dashed: ['0.5pt', 'dashed'],
  dotted: ['0.5pt', 'dotted'],
  double: ['1pt', 'double'],
};

function borderToOdf(edge: BorderEdge): string {
  const [width, style] = BORDER_TO_ODF[edge.style];
  return `${width} ${style} ${Argb.toHex(edge.color)}`;
}

function parseBorders(cellProps: XmlElement, base: Border): Border {
  const all = cellProps.attr(Fo('border'));
  const edge = (side: string, inherited: BorderEdge): BorderEdge => {
    const value = cellProps.attr(Fo(side)) ?? all;
    return value === null ? inherited : parseEdge(value);
  };
  return {
    left: edge('border-left', base.left),
    right: edge('border-right', base.right),
    top: edge('border-top', base.top),
    bottom: edge('border-bottom', base.bottom),
  };
}

/** `svg:font-family` may quote a family with spaces: `'Liberation Sans'`. */
export function unquoteFamily(family: string): string {
  const t = family.trim();
  return t.length >= 2 && (t[0] === "'" || t[0] === '"') && t.at(-1) === t[0] ? t.slice(1, -1) : t;
}

function parseEdge(value: string | null): BorderEdge {
  if (value === null || !value.trim() || value === 'none') return BorderEdges.none;

  let color: ArgbColor = Argb.black;
  let lineStyle = 'solid';
  let widthPt = 0.5;
  for (const part of value.split(' ').filter((p) => p.length > 0)) {
    const c = Argb.tryFromHex(part);
    if (c !== null) {
      color = c;
      continue;
    }
    const w = part.toLowerCase().endsWith('pt') ? parseInvariantFloat(part.slice(0, -2)) : null;
    if (w !== null) widthPt = w;
    else lineStyle = part.toLowerCase();
  }

  let style: BorderLineStyle;
  switch (lineStyle) {
    case 'dashed':
    case 'dotted':
    case 'double':
      style = lineStyle;
      break;
    default:
      style = widthPt >= 2 ? 'thick' : widthPt >= 1 ? 'medium' : 'thin';
  }
  return BorderEdges.of(style, color);
}

function parsePt(value: string | null): number | null {
  if (value === null || !value.toLowerCase().endsWith('pt')) return null;
  return parseInvariantFloat(value.slice(0, -2));
}

/** `double.TryParse(…, NumberStyles.Float, Invariant)`: optional sign, digits, point, exponent; padding allowed. */
export function parseInvariantFloat(text: string | null | undefined): number | null {
  if (text === null || text === undefined) return null;
  const t = text.trim();
  if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) return null;
  return Number(t);
}
