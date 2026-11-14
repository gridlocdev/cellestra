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
  styleKey,
  type CellStyle,
  type CellStyleMap,
  type CellStyleSource,
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

/**
 * The automatic `style:style` a cell style was read from, and the style the model read from it.
 * A save writes the element back; an edited style changes only the properties that differ from
 * `read`, so a parent style, diagonal borders, padding, indents and exact border widths survive.
 */
export interface OdsCellStyleSource extends CellStyleSource {
  readonly element: XmlElement;
  /** `styleKey` of the style as read, without its source. */
  readonly readKey: string;
  readonly read: CellStyle;
}

let nextSourceId = 1;

export function isOdsCellStyleSource(source: CellStyleSource | undefined): source is OdsCellStyleSource {
  return source !== undefined && 'element' in source && 'readKey' in source;
}

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
    for (const map of maps) element.add(mapElement(map));

    return element;
  },

  /** `style` as read from `element`, tagged with it as its source. */
  withSource(style: CellStyle, element: XmlElement): CellStyle {
    const source: OdsCellStyleSource = {
      id: nextSourceId++,
      element: element.clone(),
      readKey: styleKey(style),
      read: style,
    };
    return { ...style, source };
  },

  /**
   * Write `style` from its source element: the element as it was, renamed, with this save's data
   * style and maps; when the model style was edited since it was read, the edited properties are
   * written over the element's own and everything else in it stays.
   */
  rebuild(
    source: OdsCellStyleSource,
    name: string,
    style: CellStyle,
    dataStyleName: string | null,
    maps: readonly WrittenStyleMap[] = [],
  ): XmlElement {
    const element = source.element.clone();
    element.setAttr(Style('name'), name).setAttr(Style('family'), 'table-cell');
    element.setAttr(Style('data-style-name'), dataStyleName);
    for (const map of element.elements(Style('map'))) element.remove(map);

    const { source: _source, ...now } = style;
    if (styleKey(now) !== source.readKey) {
      applyEdits(element, source.read, now, OdsCellStyles.build(name, now, null));
    }

    for (const map of maps) element.add(mapElement(map));
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

function mapElement(map: WrittenStyleMap): XmlElement {
  const element = new XmlElement(Style('map'))
    .setAttr(Style('condition'), map.condition)
    .setAttr(Style('apply-style-name'), map.applyStyleName);
  if (map.baseCellAddress !== null) element.setAttr(Style('base-cell-address'), map.baseCellAddress);
  return element;
}

// ---- Writing an edit over a source style -----------------------------------------------------

type PropertiesKind = 'table-cell-properties' | 'paragraph-properties' | 'text-properties';
const PROPERTIES_ORDER: readonly PropertiesKind[] = [
  'table-cell-properties',
  'paragraph-properties',
  'text-properties',
];

/**
 * One modeled property and the attributes that spell it. When the property changed, every one of
 * `owned` is removed and `written` is set from the generated style, or to `off` when the generated
 * style leaves it out (the source may have a parent that sets it, so leaving it out is not "off").
 * `mirrors` are the Asian and complex-script twins LibreOffice writes beside a font property; when
 * the source had one, it takes the new value too.
 */
interface PropertyRule {
  readonly changed: (a: CellStyle, b: CellStyle) => boolean;
  readonly kind: PropertiesKind;
  readonly owned: readonly XmlName[];
  readonly written: readonly XmlName[];
  readonly off: string | null;
  readonly mirrors?: readonly XmlName[];
}

const differs =
  <T>(get: (s: CellStyle) => T) =>
  (a: CellStyle, b: CellStyle): boolean =>
    get(a) !== get(b);

const PROPERTY_RULES: readonly PropertyRule[] = [
  {
    changed: differs((s) => s.font.bold),
    kind: 'text-properties',
    owned: [Fo('font-weight')],
    written: [Fo('font-weight')],
    off: 'normal',
    mirrors: [Style('font-weight-asian'), Style('font-weight-complex')],
  },
  {
    changed: differs((s) => s.font.italic),
    kind: 'text-properties',
    owned: [Fo('font-style')],
    written: [Fo('font-style')],
    off: 'normal',
    mirrors: [Style('font-style-asian'), Style('font-style-complex')],
  },
  {
    changed: differs((s) => s.font.underline),
    kind: 'text-properties',
    owned: ['style', 'type', 'width', 'color', 'mode'].map((p) => Style(`text-underline-${p}`)),
    written: [Style('text-underline-style'), Style('text-underline-width'), Style('text-underline-color')],
    off: 'none',
  },
  {
    changed: differs((s) => s.font.strikethrough),
    kind: 'text-properties',
    owned: ['style', 'type', 'width', 'color', 'mode', 'text', 'text-style'].map((p) =>
      Style(`text-line-through-${p}`),
    ),
    written: [Style('text-line-through-style')],
    off: 'none',
  },
  {
    changed: differs((s) => s.font.color >>> 0),
    kind: 'text-properties',
    owned: [Fo('color')],
    written: [Fo('color')],
    off: '#000000',
  },
  {
    changed: differs((s) => s.font.family),
    kind: 'text-properties',
    owned: [Fo('font-family'), Style('font-name'), Style('font-family-generic'), Style('font-pitch')],
    written: [Fo('font-family')],
    off: Defaults.fontFamily,
  },
  {
    changed: differs((s) => s.font.sizePoints),
    kind: 'text-properties',
    owned: [Fo('font-size')],
    written: [Fo('font-size')],
    off: `${Defaults.fontSizePoints}pt`,
    mirrors: [Style('font-size-asian'), Style('font-size-complex')],
  },
  {
    changed: differs((s) => (s.fill.background === null ? null : s.fill.background >>> 0)),
    kind: 'table-cell-properties',
    owned: [Fo('background-color')],
    written: [Fo('background-color')],
    off: 'transparent',
  },
  {
    changed: differs((s) => s.alignment.vertical),
    kind: 'table-cell-properties',
    owned: [Style('vertical-align')],
    written: [Style('vertical-align')],
    off: null,
  },
  {
    changed: differs((s) => s.alignment.wrapText),
    kind: 'table-cell-properties',
    owned: [Fo('wrap-option')],
    written: [Fo('wrap-option')],
    off: 'no-wrap',
  },
  {
    changed: (a, b) =>
      (a.alignment.rotation ?? 0) !== (b.alignment.rotation ?? 0) ||
      a.alignment.rotationAlign !== b.alignment.rotationAlign,
    kind: 'table-cell-properties',
    owned: [Style('rotation-angle'), Style('rotation-align')],
    written: [Style('rotation-angle'), Style('rotation-align')],
    off: '0',
  },
  {
    changed: (a, b) => {
      const p = a.protection ?? Protections.default;
      const q = b.protection ?? Protections.default;
      return (
        p.locked !== q.locked ||
        p.hideFormula !== q.hideFormula ||
        p.hideAll !== q.hideAll ||
        p.printContent !== q.printContent
      );
    },
    kind: 'table-cell-properties',
    owned: [Style('cell-protect'), Style('print-content')],
    written: [Style('cell-protect'), Style('print-content')],
    off: 'protected',
  },
];

function applyEdits(element: XmlElement, read: CellStyle, now: CellStyle, generated: XmlElement): void {
  const generatedProps = (kind: PropertiesKind) => generated.element(Style(kind));

  for (const rule of PROPERTY_RULES) {
    if (!rule.changed(read, now)) continue;
    const props = propertiesOf(element, rule.kind);
    const hadMirror = (rule.mirrors ?? []).filter((m) => props.attr(m) !== null);
    for (const name of [...rule.owned, ...(rule.mirrors ?? [])]) props.setAttr(name, null);
    const source = generatedProps(rule.kind);
    let value: string | null = null;
    for (const name of rule.written) {
      const v = source?.attr(name) ?? null;
      if (v !== null) props.setAttr(name, v);
      if (name === rule.written[0]) value = v;
    }
    if (value === null && rule.off !== null) props.setAttr(rule.written[0]!, (value = rule.off));
    if (value !== null) for (const mirror of hadMirror) props.setAttr(mirror, value);
  }

  applyBorderEdits(element, read.border, now.border, generatedProps('table-cell-properties'));

  if (read.alignment.horizontal !== now.alignment.horizontal) {
    const paragraph = propertiesOf(element, 'paragraph-properties');
    const cell = propertiesOf(element, 'table-cell-properties');
    const fixed = now.alignment.horizontal !== 'general';
    paragraph.setAttr(Fo('text-align'), fixed ? HORIZONTAL_TO_ODF[now.alignment.horizontal] : null);
    cell.setAttr(Style('text-align-source'), fixed ? 'fix' : 'value-type');
  }

  for (const kind of PROPERTIES_ORDER) {
    const props = element.element(Style(kind));
    if (props && !props.hasAttributes && !props.hasElements) element.remove(props);
  }
}

const SIDES = ['left', 'right', 'top', 'bottom'] as const;

/**
 * Write the edited border edges. A `fo:border` shorthand is spelt out per side first, so the sides
 * that did not change keep their exact width and style.
 */
function applyBorderEdits(
  element: XmlElement,
  read: Border,
  now: Border,
  generated: XmlElement | null,
): void {
  const key = (e: BorderEdge) => `${e.style}:${e.color >>> 0}`;
  const changed = SIDES.filter((side) => key(read[side]) !== key(now[side]));
  if (changed.length === 0) return;
  const props = propertiesOf(element, 'table-cell-properties');
  const all = props.attr(Fo('border'));
  const allWidth = props.attr(Style('border-line-width'));
  props.setAttr(Fo('border'), null).setAttr(Style('border-line-width'), null);
  for (const side of SIDES) {
    if (changed.includes(side)) {
      props.setAttr(Style(`border-line-width-${side}`), null);
      props.setAttr(Fo(`border-${side}`), generated?.attr(Fo(`border-${side}`)) ?? 'none');
      continue;
    }
    if (all !== null && props.attr(Fo(`border-${side}`)) === null) props.setAttr(Fo(`border-${side}`), all);
    if (allWidth !== null && props.attr(Style(`border-line-width-${side}`)) === null)
      props.setAttr(Style(`border-line-width-${side}`), allWidth);
  }
}

/** The element's `style:<kind>` child, added in schema order when it has none. */
function propertiesOf(element: XmlElement, kind: PropertiesKind): XmlElement {
  const existing = element.element(Style(kind));
  if (existing) return existing;
  const created = new XmlElement(Style(kind));
  const later = PROPERTIES_ORDER.slice(PROPERTIES_ORDER.indexOf(kind) + 1);
  const next = element.elements().find((e) => e.is(Style('map')) || later.some((k) => e.is(Style(k))));
  if (!next) {
    element.add(created);
  } else {
    created.parent = element;
    element.nodes.splice(element.nodes.indexOf(next), 0, created);
  }
  return created;
}

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
