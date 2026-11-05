import { Defaults } from '../../../defaults.ts';
import {
  Alignments,
  BorderEdges,
  Borders,
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
  type NumberFormat,
  type VerticalAlignment,
} from '../../../styling/cellStyle.ts';
import { Argb, type ArgbColor } from '../../../styling/color.ts';
import { XmlElement, type XmlName } from '../../../xml/index.ts';
import { Fo, Style } from '../odsXmlNames.ts';

/** Maps a `CellStyle` to an ODF table-cell automatic style and back. */
export const OdsCellStyles = {
  build(name: string, style: CellStyle, dataStyleName: string | null): XmlElement {
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

    return element;
  },

  parse(styleElement: XmlElement, resolveDataStyle: (name: string) => NumberFormat): CellStyle {
    let font: Font = Fonts.default;
    let fill: Fill = Fills.none;
    let alignment: Alignment = Alignments.default;
    let border: Border = Borders.none;

    const textProps = styleElement.element(Style('text-properties'));
    if (textProps) {
      const underline = textProps.attr(Style('text-underline-style'));
      const strike = textProps.attr(Style('text-line-through-style'));
      font = {
        ...font,
        bold: textProps.attr(Fo('font-weight')) === 'bold',
        italic: textProps.attr(Fo('font-style')) === 'italic',
        underline: underline !== null && underline !== 'none',
        strikethrough: strike !== null && strike !== 'none',
      };
      const color = Argb.tryFromHex(textProps.attr(Fo('color')));
      if (color !== null) font = { ...font, color };
      const family = textProps.attr(Fo('font-family'));
      if (family) font = { ...font, family };
      const size = parsePt(textProps.attr(Fo('font-size')));
      if (size !== null) font = { ...font, sizePoints: size };
    }

    const cellProps = styleElement.element(Style('table-cell-properties'));
    if (cellProps) {
      const background = cellProps.attr(Fo('background-color'));
      const bg = Argb.tryFromHex(background);
      if (bg !== null && background !== 'transparent') fill = Fills.solid(bg);

      const vertical: VerticalAlignment =
        VERTICAL_FROM_ODF[cellProps.attr(Style('vertical-align')) ?? ''] ?? 'bottom';
      const wrap = cellProps.attr(Fo('wrap-option')) === 'wrap';
      alignment = { ...alignment, vertical, wrapText: wrap };
      border = parseBorders(cellProps);
    }

    const align = styleElement.element(Style('paragraph-properties'))?.attr(Fo('text-align')) ?? null;
    if (align !== null) {
      alignment = { ...alignment, horizontal: HORIZONTAL_FROM_ODF[align] ?? 'general' };
    }

    const dataStyleName = styleElement.attr(Style('data-style-name'));
    const numberFormat = dataStyleName === null ? NumberFormats.general : resolveDataStyle(dataStyleName);

    return { font, fill, border, alignment, numberFormat };
  },
} as const;

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

function parseBorders(cellProps: XmlElement): Border {
  const all = cellProps.attr(Fo('border'));
  const edge = (side: string): BorderEdge => parseEdge(cellProps.attr(Fo(side)) ?? all);
  return {
    left: edge('border-left'),
    right: edge('border-right'),
    top: edge('border-top'),
    bottom: edge('border-bottom'),
  };
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
