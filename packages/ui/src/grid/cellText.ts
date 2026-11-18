import type { CellStyle } from '@cellestra/core';

/**
 * How a cell's text is sized: its font, its padding, its line height, the box turned text needs.
 * The grid draws with these, and row heights and column autofit measure with them, so a fitted
 * column or a grown row is exactly what the grid needs to show the text whole.
 */

/** Gap between a cell's text and its left and right edges, in CSS px at 100%. */
export const CELL_PAD_X = 5;
/** Gap between a cell's text and its top and bottom edges, in CSS px at 100%. */
export const CELL_PAD_Y = 2;
export const PT_TO_PX = 96 / 72;
/**
 * A line of text is about 1.2× its font size. With the 2px padding above and below, a 10pt
 * line comes to exactly the default 20px row, so only bigger fonts and wrapped text grow a row.
 */
export const LINE_FACTOR = 1.2;

/** What measuring needs from a canvas context. */
export interface TextMeasure {
  font: string;
  measureText(text: string): { readonly width: number };
}

/** The CSS font family list for a cell font: its own family, then fallbacks. */
export function cssFamily(family: string): string {
  return `"${family.replace(/"/g, '')}", "Helvetica Neue", Helvetica, Arial, sans-serif`;
}

/** The CSS font a cell's text is drawn and measured in, at `zoom`. */
export function cellFont(font: CellStyle['font'], zoom = 1): string {
  const px = font.sizePoints * PT_TO_PX * zoom;
  return `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${px}px ${cssFamily(font.family)}`;
}

/** The height one cell's text needs: `lines` lines at `points`, plus the padding. */
export function textHeight(points: number, lines = 1): number {
  return Math.ceil(points * PT_TO_PX * LINE_FACTOR * Math.max(1, lines) + 2 * CELL_PAD_Y);
}

/** The upright bounding box of a `w` × `h` block turned `degrees`. */
export function rotatedBox(w: number, h: number, degrees: number): { width: number; height: number } {
  const t = (degrees * Math.PI) / 180;
  const c = Math.abs(Math.cos(t));
  const s = Math.abs(Math.sin(t));
  return { width: w * c + h * s, height: w * s + h * c };
}

/**
 * The upright room a cell's turned text needs, padding included: the text's `width` × `height` with
 * the cell padding added along it (so the ends get CELL_PAD_X), then turned `degrees`.
 */
export function rotatedCellBox(
  width: number,
  height: number,
  degrees: number,
): { width: number; height: number } {
  return rotatedBox(width + 2 * CELL_PAD_X, height + 2 * CELL_PAD_Y, degrees);
}

/** The width of `text` in the context's current font: a multi-line value is as wide as its longest line. */
export function widestLine(ctx: TextMeasure, text: string): number {
  return Math.max(...text.split('\n').map((line) => ctx.measureText(line).width));
}

/**
 * The column width, at 100%, that shows `text` whole in a cell of `style`: flat text plus its
 * padding each side, turned text its upright box. Wrapped flat text folds to any width, so it asks
 * for none (0).
 */
export function fittedWidth(ctx: TextMeasure, text: string, style: CellStyle): number {
  const rotation = style.alignment.rotation ?? 0;
  if (style.alignment.wrapText && rotation === 0) return 0;
  ctx.font = cellFont(style.font);
  const width = widestLine(ctx, text);
  if (rotation === 0) return width + 2 * CELL_PAD_X;
  const lines = text.split('\n').length;
  const height = style.font.sizePoints * PT_TO_PX * LINE_FACTOR * lines;
  return rotatedCellBox(width, height, rotation).width;
}

/** `text` cut short with an ellipsis so it measures at most `width` (empty when nothing fits). */
export function fitText(ctx: TextMeasure, text: string, width: number): string {
  if (ctx.measureText(text).width <= width) return text;
  const chars = [...text];
  // The longest prefix that fits with its ellipsis (binary search: titles can be long).
  let lo = 0;
  let hi = chars.length - 1;
  while (lo < hi) {
    const mid = Math.ceil((lo + hi) / 2);
    if (ctx.measureText(chars.slice(0, mid).join('') + '…').width <= width) lo = mid;
    else hi = mid - 1;
  }
  return lo > 0 ? chars.slice(0, lo).join('') + '…' : '';
}
