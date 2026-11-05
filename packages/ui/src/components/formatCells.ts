import {
  Argb,
  BorderEdges,
  CellValue,
  Fills,
  NumberFormatter,
  NumberFormats,
  type ArgbColor,
  type Border,
  type BorderEdge,
  type BorderLineStyle,
  type CellStyle,
  type HorizontalAlignment,
  type VerticalAlignment,
} from '@cellestra/core';

/**
 * Format Cells dialog logic (PRD §6.2, ⌘1), kept free of Preact so it is testable.
 *
 * The dialog seeds from the active cell's style and the distinct styles across the selection. A field
 * on which the selection disagrees seeds to "varies" (`null`). Every edit marks its field touched, and
 * the transform rewrites touched fields only, so a mixed selection keeps each cell's own value for
 * every field the user left alone.
 */

export const VARIES = '(varies)';

/** The preset number formats, in menu order. */
export const NUMBER_FORMAT_PRESETS: readonly { readonly label: string; readonly code: string }[] = [
  { label: 'General', code: 'General' },
  { label: 'Number (0)', code: '0' },
  { label: 'Number (0.00)', code: '0.00' },
  { label: 'Thousands (#,##0)', code: '#,##0' },
  { label: 'Thousands (#,##0.00)', code: '#,##0.00' },
  { label: 'Percent (0%)', code: '0%' },
  { label: 'Percent (0.00%)', code: '0.00%' },
  { label: 'Currency ($#,##0.00)', code: '$#,##0.00' },
  { label: 'Date (yyyy-mm-dd)', code: 'yyyy-mm-dd' },
];

export const BORDER_LINE_STYLES: readonly { readonly label: string; readonly style: BorderLineStyle }[] = [
  { label: 'Thin', style: 'thin' },
  { label: 'Medium', style: 'medium' },
  { label: 'Thick', style: 'thick' },
  { label: 'Dashed', style: 'dashed' },
  { label: 'Dotted', style: 'dotted' },
  { label: 'Double', style: 'double' },
];

export const HORIZONTAL_ALIGNMENTS: readonly {
  readonly label: string;
  readonly value: HorizontalAlignment;
}[] = [
  { label: 'General', value: 'general' },
  { label: 'Left', value: 'left' },
  { label: 'Center', value: 'center' },
  { label: 'Right', value: 'right' },
];

export const VERTICAL_ALIGNMENTS: readonly { readonly label: string; readonly value: VerticalAlignment }[] = [
  { label: 'Top', value: 'top' },
  { label: 'Middle', value: 'middle' },
  { label: 'Bottom', value: 'bottom' },
];

export const FONT_SIZE_MIN = 6;
export const FONT_SIZE_MAX = 96;

export interface BorderEdgesOn {
  readonly top: boolean;
  readonly bottom: boolean;
  readonly left: boolean;
  readonly right: boolean;
}

/** Border presets. "Outline" equals "All" because the transform runs per cell (as in the desktop app). */
export const BORDER_PRESETS: readonly { readonly label: string; readonly edges: BorderEdgesOn }[] = [
  { label: 'None', edges: { top: false, bottom: false, left: false, right: false } },
  { label: 'Outline', edges: { top: true, bottom: true, left: true, right: true } },
  { label: 'All', edges: { top: true, bottom: true, left: true, right: true } },
  { label: 'Top', edges: { top: true, bottom: false, left: false, right: false } },
  { label: 'Bottom', edges: { top: false, bottom: true, left: false, right: false } },
  { label: 'Left', edges: { top: false, bottom: false, left: true, right: false } },
  { label: 'Right', edges: { top: false, bottom: false, left: false, right: true } },
];

/** The dialog's values. `null` means "(varies)": the selection disagrees and the user has not chosen. */
export interface FormatCellsValues {
  readonly numberFormat: string | null;
  readonly bold: boolean | null;
  readonly italic: boolean | null;
  readonly underline: boolean | null;
  readonly strikethrough: boolean | null;
  readonly size: number | null;
  readonly textColor: ArgbColor;
  readonly fill: boolean | null;
  readonly fillColor: ArgbColor;
  readonly horizontal: HorizontalAlignment | null;
  readonly vertical: VerticalAlignment | null;
  readonly wrapText: boolean | null;
  readonly borderStyle: BorderLineStyle;
  readonly borderColor: ArgbColor;
  readonly borderEdges: BorderEdgesOn;
}

export type FormatCellsField = keyof FormatCellsValues;

/** Which fields vary across the selection. Colours and the border can't show "varies", so they are listed here. */
export type FormatCellsVaries = { readonly [K in FormatCellsField]: boolean };

export interface FormatCellsState {
  readonly values: FormatCellsValues;
  readonly varies: FormatCellsVaries;
  readonly touched: ReadonlySet<FormatCellsField>;
}

/** True when `select` gives the same value for every style (by `===`, or `equals` when given). */
export function allSame<T>(
  styles: readonly CellStyle[],
  select: (s: CellStyle) => T,
  equals: (a: T, b: T) => boolean = (a, b) => a === b,
): boolean {
  const first = styles[0];
  if (first === undefined) return true;
  const seed = select(first);
  return styles.every((s) => equals(seed, select(s)));
}

const sameEdge = (a: BorderEdge, b: BorderEdge) => a.style === b.style && a.color >>> 0 === b.color >>> 0;
const sameBorder = (a: Border, b: Border) =>
  sameEdge(a.left, b.left) &&
  sameEdge(a.right, b.right) &&
  sameEdge(a.top, b.top) &&
  sameEdge(a.bottom, b.bottom);
const sameColor = (a: ArgbColor | null, b: ArgbColor | null) =>
  a === null || b === null ? a === b : a >>> 0 === b >>> 0;

/** The first visible edge, to seed the line style and colour; a thin black edge when there is none. */
export function firstVisibleEdge(border: Border): BorderEdge {
  for (const e of [border.top, border.bottom, border.left, border.right]) {
    if (BorderEdges.isVisible(e)) return e;
  }
  return BorderEdges.of('thin');
}

/** Seed the dialog from the active cell's style and the selection's distinct styles. */
export function seedFormatCells(current: CellStyle, styles: readonly CellStyle[]): FormatCellsState {
  const all = styles.length > 0 ? styles : [current];
  const font = current.font;
  const edge = firstVisibleEdge(current.border);
  const varies: FormatCellsVaries = {
    numberFormat: !allSame(all, (s) => s.numberFormat.code),
    bold: !allSame(all, (s) => s.font.bold),
    italic: !allSame(all, (s) => s.font.italic),
    underline: !allSame(all, (s) => s.font.underline),
    strikethrough: !allSame(all, (s) => s.font.strikethrough),
    size: !allSame(all, (s) => s.font.sizePoints),
    textColor: !allSame(all, (s) => s.font.color, sameColor),
    fill: !allSame(all, (s) => s.fill.background, sameColor),
    fillColor: !allSame(all, (s) => s.fill.background, sameColor),
    horizontal: !allSame(all, (s) => s.alignment.horizontal),
    vertical: !allSame(all, (s) => s.alignment.vertical),
    wrapText: !allSame(all, (s) => s.alignment.wrapText),
    borderStyle: !allSame(all, (s) => s.border, sameBorder),
    borderColor: !allSame(all, (s) => s.border, sameBorder),
    borderEdges: !allSame(all, (s) => s.border, sameBorder),
  };
  const pick = <T>(field: FormatCellsField, value: T): T | null => (varies[field] ? null : value);
  const values: FormatCellsValues = {
    numberFormat: pick('numberFormat', current.numberFormat.code),
    bold: pick('bold', font.bold),
    italic: pick('italic', font.italic),
    underline: pick('underline', font.underline),
    strikethrough: pick('strikethrough', font.strikethrough),
    size: pick('size', font.sizePoints),
    textColor: Argb.isTransparent(font.color) ? Argb.black : font.color,
    fill: pick('fill', Fills.hasBackground(current.fill)),
    fillColor: current.fill.background ?? Argb.white,
    horizontal: pick('horizontal', current.alignment.horizontal),
    vertical: pick('vertical', current.alignment.vertical),
    wrapText: pick('wrapText', current.alignment.wrapText),
    borderStyle: BorderEdges.isVisible(edge) ? edge.style : 'thin',
    borderColor: Argb.isTransparent(edge.color) ? Argb.black : edge.color,
    borderEdges: {
      top: BorderEdges.isVisible(current.border.top),
      bottom: BorderEdges.isVisible(current.border.bottom),
      left: BorderEdges.isVisible(current.border.left),
      right: BorderEdges.isVisible(current.border.right),
    },
  };
  return { values, varies, touched: new Set() };
}

/** Set fields and mark them touched. Setting a varying field back to `null` un-touches it. */
export function editFormatCells(
  state: FormatCellsState,
  patch: Partial<FormatCellsValues>,
): FormatCellsState {
  const touched = new Set(state.touched);
  for (const key of Object.keys(patch) as FormatCellsField[]) {
    if (patch[key] === null) touched.delete(key);
    else touched.add(key);
  }
  return { ...state, values: { ...state.values, ...patch }, touched };
}

/** A tri-state checkbox click: unchecked → checked → (varies, when the field varies) → unchecked. */
export function nextTriState(value: boolean | null, allowVaries: boolean): boolean | null {
  if (value === null) return true;
  if (value) return false;
  return allowVaries ? null : true;
}

/** Clamp a typed font size; null when it is not a number. */
export function parseFontSize(text: string): number | null {
  const n = Number.parseFloat(text);
  if (!Number.isFinite(n)) return null;
  return Math.min(FONT_SIZE_MAX, Math.max(FONT_SIZE_MIN, n));
}

/** The preset index for a code, or -1 when it is a custom code. */
export function presetIndex(code: string): number {
  const normal = NumberFormats.isGeneral({ code }) ? 'General' : code;
  return NUMBER_FORMAT_PRESETS.findIndex((p) => p.code === normal);
}

const SAMPLE_NUMBER = 1234.5678;
const SAMPLE_NEGATIVE = -1234.5678;
/** 2025-03-14 15:30 as a spreadsheet serial (1899-12-30 epoch). */
const SAMPLE_DATE = 45730.645833333336;

/** The sample value a code is previewed with: a date serial for date/time codes, text for `@`. */
export function sampleValueFor(code: string): CellValue {
  switch (NumberFormats.kind({ code })) {
    case 'date':
    case 'time':
    case 'dateTime':
      return CellValue.number(SAMPLE_DATE);
    case 'text':
      return CellValue.text('Sample text');
    case 'percent':
      return CellValue.number(0.1234);
    default:
      return CellValue.number(SAMPLE_NUMBER);
  }
}

/** A live preview of a code: the sample, and a negative sample for numeric kinds. Never throws. */
export function formatSample(
  code: string,
  locale: string,
): { readonly text: string; readonly negative?: string } {
  try {
    const kind = NumberFormats.kind({ code });
    const value = sampleValueFor(code);
    const text = NumberFormatter.format(value, code, locale);
    if (kind === 'number' || kind === 'currency' || kind === 'general' || kind === 'scientific') {
      return { text, negative: NumberFormatter.format(CellValue.number(SAMPLE_NEGATIVE), code, locale) };
    }
    return { text };
  } catch {
    return { text: '' };
  }
}

/** The border the dialog would write: the chosen edge on every checked side, none elsewhere. */
export function dialogBorder(values: FormatCellsValues): Border {
  const edge = BorderEdges.of(values.borderStyle, values.borderColor);
  const on = values.borderEdges;
  return {
    left: on.left ? edge : BorderEdges.none,
    right: on.right ? edge : BorderEdges.none,
    top: on.top ? edge : BorderEdges.none,
    bottom: on.bottom ? edge : BorderEdges.none,
  };
}

/** The style transform for every selected cell: only touched fields are rewritten. */
export function buildFormatTransform(state: FormatCellsState): (style: CellStyle) => CellStyle {
  const { values: v, touched } = state;
  const has = (f: FormatCellsField) => touched.has(f);
  const code =
    has('numberFormat') && v.numberFormat !== null
      ? v.numberFormat.trim() || NumberFormats.general.code
      : null;
  const tri = (f: 'bold' | 'italic' | 'underline' | 'strikethrough') => (has(f) ? v[f] : null);
  const bold = tri('bold');
  const italic = tri('italic');
  const underline = tri('underline');
  const strikethrough = tri('strikethrough');
  const size = has('size') ? v.size : null;
  const textColor = has('textColor') ? v.textColor : null;
  // Fill: off → none; on → solid, in the picked colour when the colour was touched, else each cell's own.
  const fillOn = has('fill') || has('fillColor') ? v.fill : null;
  const fillColor = has('fillColor') ? v.fillColor : null;
  const horizontal = has('horizontal') ? v.horizontal : null;
  const vertical = has('vertical') ? v.vertical : null;
  const wrapText = has('wrapText') ? v.wrapText : null;
  const border = has('borderStyle') || has('borderColor') || has('borderEdges') ? dialogBorder(v) : null;

  return (style) => {
    let font = style.font;
    if (bold !== null) font = { ...font, bold };
    if (italic !== null) font = { ...font, italic };
    if (underline !== null) font = { ...font, underline };
    if (strikethrough !== null) font = { ...font, strikethrough };
    if (size !== null) font = { ...font, sizePoints: size };
    if (textColor !== null) font = { ...font, color: textColor };

    let fill = style.fill;
    if (fillOn === false) fill = Fills.none;
    else if (fillOn === true) fill = Fills.solid(fillColor ?? style.fill.background ?? v.fillColor);

    let alignment = style.alignment;
    if (horizontal !== null) alignment = { ...alignment, horizontal };
    if (vertical !== null) alignment = { ...alignment, vertical };
    if (wrapText !== null) alignment = { ...alignment, wrapText };

    const changed =
      font !== style.font ||
      fill !== style.fill ||
      alignment !== style.alignment ||
      border !== null ||
      code !== null;
    if (!changed) return style;
    return {
      font,
      fill,
      alignment,
      border: border ?? style.border,
      numberFormat: code !== null ? NumberFormats.of(code) : style.numberFormat,
    };
  };
}
