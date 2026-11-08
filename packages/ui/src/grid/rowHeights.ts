import { Defaults, StylePool, type CellStyle, type CellValue, type Worksheet } from '@cellestra/core';
import { CELL_PAD_Y } from './render.ts';

const PT_TO_PX = 96 / 72;
/**
 * A line of text is about 1.2× its font size. With the 2px padding above and below, a 10pt
 * line comes to exactly the default 20px row, so only bigger fonts and wrapped text grow a row.
 */
const LINE_FACTOR = 1.2;

/** The height one cell's text needs: `lines` lines at `points`, plus the padding. */
export function textHeight(points: number, lines = 1): number {
  return Math.ceil(points * PT_TO_PX * LINE_FACTOR * Math.max(1, lines) + 2 * CELL_PAD_Y);
}

/** What the measuring needs from the grid: the shown text, and how many lines wrapping makes of it. */
export interface RowMeasure {
  display(value: CellValue, style: CellStyle): string;
  /** Lines `text` wraps to in `width` px with `style`'s font. */
  wrappedLines(text: string, style: CellStyle, width: number): number;
  columnWidth(column: number): number;
}

/**
 * The height each row's content needs, for rows where that is more than the default: the tallest
 * font in the row, and the tallest wrapped cell. The grid draws a row at the larger of this and its
 * own height, so a bigger font or wrapped text never overlaps the next row, and a row resized by
 * hand to be taller keeps its size. One pass over the stored cells.
 */
export function contentRowHeights(sheet: Worksheet, styles: StylePool, m: RowMeasure): Map<number, number> {
  const out = new Map<number, number>();
  const single = new Map<number, number>();
  for (const [row, column, cell] of sheet.enumerateCells()) {
    if (cell.value.kind === 'empty') continue;
    const id = sheet.effectiveStyleId(row, column);
    // Line breaks typed with ⌥↩ show as separate lines whether or not the cell wraps.
    const breaks = cell.value.kind === 'text' ? cell.value.value.split('\n').length : 1;
    if (id === StylePool.defaultId && breaks === 1) continue;
    // A merged cell's text has the merge's whole height to sit in; leave those rows alone.
    if (sheet.findMerge(row, column)) continue;
    const style = styles.get(id);
    let need = single.get(id);
    if (need === undefined) {
      need = textHeight(style.font.sizePoints);
      single.set(id, need);
    }
    if (breaks > 1 && !style.alignment.wrapText)
      need = Math.max(need, textHeight(style.font.sizePoints, breaks));
    if (style.alignment.wrapText) {
      const text = m.display(cell.value, style);
      const lines = m.wrappedLines(text, style, m.columnWidth(column) - 2 * 5);
      need = Math.max(need, textHeight(style.font.sizePoints, lines));
    }
    if (need > Defaults.rowHeight && need > (out.get(row) ?? 0)) out.set(row, need);
  }
  return out;
}
