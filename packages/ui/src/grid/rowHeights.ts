import { Defaults, StylePool, type CellStyle, type CellValue, type Worksheet } from '@cellestra/core';
import { CELL_PAD_Y, rotatedBox } from './render.ts';

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
  /** Width of one line of `text` in `style`'s font; rotated text needs it to size its row. */
  textWidth?(text: string, style: CellStyle): number;
}

/**
 * The height each row's content needs, for rows where that is more than the default: the tallest
 * font in the row, and the tallest wrapped cell. The grid draws a row at the larger of this and its
 * own height, so a bigger font or wrapped text never overlaps the next row, and a row resized by
 * hand to be taller keeps its size.
 *
 * Built with one pass over the stored cells; after an edit, `update` re-measures only the cells
 * the edit touched, so typing in a big sheet doesn't walk every cell again.
 */
export class ContentHeights {
  /** Row → the height its content needs, only where that is more than the default. */
  readonly rows = new Map<number, number>();
  /** Row → column → that cell's need, for the cells that make their row taller. */
  readonly #cells = new Map<number, Map<number, number>>();
  readonly #single = new Map<number, number>();
  readonly sheet: Worksheet;
  readonly #styles: StylePool;
  readonly #m: RowMeasure;

  constructor(sheet: Worksheet, styles: StylePool, m: RowMeasure) {
    this.sheet = sheet;
    this.#styles = styles;
    this.#m = m;
    // Rows some merge covers: only cells in those rows need the (linear) merge lookup.
    const mergedRows = new Set<number>();
    for (const merge of sheet.merges)
      for (let r = merge.range.top; r <= merge.range.bottom && mergedRows.size < 100_000; r++)
        mergedRows.add(r);
    const capped = mergedRows.size >= 100_000;
    for (const [row, column, cell] of sheet.enumerateCells()) {
      if (cell.value.kind === 'empty') continue;
      const need = this.#need(row, column, cell.value, capped || mergedRows.has(row));
      if (need > Defaults.rowHeight) this.#set(row, column, need);
    }
    for (const row of this.#cells.keys()) this.#refresh(row);
  }

  /** Re-measure these cells after an edit. True when some row's height changed. */
  update(cells: Iterable<{ readonly row: number; readonly column: number }>): boolean {
    const touched = new Set<number>();
    for (const { row, column } of cells) {
      const value = this.sheet.find(row, column)?.value;
      const need = value && value.kind !== 'empty' ? this.#need(row, column, value, true) : 0;
      const had = this.#cells.get(row)?.get(column);
      if (need > Defaults.rowHeight) {
        if (had === need) continue;
        this.#set(row, column, need);
      } else {
        if (had === undefined) continue;
        const inRow = this.#cells.get(row)!;
        inRow.delete(column);
        if (inRow.size === 0) this.#cells.delete(row);
      }
      touched.add(row);
    }
    let changed = false;
    for (const row of touched) if (this.#refresh(row)) changed = true;
    return changed;
  }

  #set(row: number, column: number, need: number): void {
    let inRow = this.#cells.get(row);
    if (!inRow) this.#cells.set(row, (inRow = new Map()));
    inRow.set(column, need);
  }

  /** Recompute a row's height from its cells; true when it changed. */
  #refresh(row: number): boolean {
    let max = 0;
    for (const need of this.#cells.get(row)?.values() ?? []) if (need > max) max = need;
    const before = this.rows.get(row);
    if (max > Defaults.rowHeight) {
      if (before === max) return false;
      this.rows.set(row, max);
      return true;
    }
    if (before === undefined) return false;
    this.rows.delete(row);
    return true;
  }

  /** The height one cell needs; 0 when it doesn't grow its row. */
  #need(row: number, column: number, value: CellValue, mayBeMerged: boolean): number {
    const sheet = this.sheet;
    const id = sheet.effectiveStyleId(row, column);
    // Line breaks typed with ⌥↩ show as separate lines whether or not the cell wraps.
    const breaks = value.kind === 'text' ? value.value.split('\n').length : 1;
    if (id === StylePool.defaultId && breaks === 1) return 0;
    // A merged cell's text has the merge's whole height to sit in; leave those rows alone.
    if (mayBeMerged && sheet.findMerge(row, column)) return 0;
    const style = this.#styles.get(id);
    let need = this.#single.get(id);
    if (need === undefined) {
      need = textHeight(style.font.sizePoints);
      this.#single.set(id, need);
    }
    const rotation = style.alignment.rotation ?? 0;
    if (rotation !== 0 && this.#m.textWidth) {
      const lines = this.#m.display(value, style).split('\n');
      const width = Math.max(...lines.map((t) => this.#m.textWidth!(t, style)));
      const height = style.font.sizePoints * PT_TO_PX * LINE_FACTOR * lines.length;
      return Math.ceil(rotatedBox(width, height, rotation).height + 2 * CELL_PAD_Y);
    }
    if (breaks > 1 && !style.alignment.wrapText)
      need = Math.max(need, textHeight(style.font.sizePoints, breaks));
    if (style.alignment.wrapText) {
      const text = this.#m.display(value, style);
      const lines = this.#m.wrappedLines(text, style, this.#m.columnWidth(column) - 2 * 5);
      need = Math.max(need, textHeight(style.font.sizePoints, lines));
    }
    return need;
  }
}

/** The rows content makes taller, measured in one pass (see `ContentHeights`). */
export function contentRowHeights(sheet: Worksheet, styles: StylePool, m: RowMeasure): Map<number, number> {
  return new ContentHeights(sheet, styles, m).rows;
}
