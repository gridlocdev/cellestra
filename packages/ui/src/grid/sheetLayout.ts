import type { Worksheet } from '@cellestra/core';
import { MAX_COLUMN, MAX_ROW } from '../session/selection.ts';
import { Axis } from './axis.ts';

/** A column's width and a row's height when the sheet gives it none, in CSS px at 100%. */
const DEFAULT_COLUMN_WIDTH = 64;
const DEFAULT_ROW_HEIGHT = 20;

/**
 * Where a sheet's lines fall on screen: column widths and row heights (hidden lines take none,
 * rows grow to fit their content), and how far the scroller must reach.
 */
export interface SheetLayout {
  readonly columns: Axis;
  readonly rows: Axis;
  /** The last row and column with non-default layout (−1 for none), for the scroll extent. */
  readonly lastRow: number;
  readonly lastColumn: number;
}

/** The layout of `sheet`, with rows at least as tall as `content` (row → height its text needs). */
export function sheetLayout(sheet: Worksheet, content: ReadonlyMap<number, number>): SheetLayout {
  const cols: [number, number][] = [];
  let lastColumn = -1;
  for (const [i, c] of sheet.columns) {
    cols.push([i, c.hidden ? 0 : c.width]);
    if (!c.isDefault) lastColumn = Math.max(lastColumn, i);
  }
  const rows: [number, number][] = [];
  let lastRow = -1;
  for (const [i, r] of sheet.rows) {
    rows.push([i, r.hidden ? 0 : Math.max(r.height, content.get(i) ?? 0)]);
    if (!r.isDefault) lastRow = Math.max(lastRow, i);
  }
  // Rows with no layout of their own that their content makes taller.
  for (const [i, h] of content) if (!sheet.rows.has(i)) rows.push([i, h]);
  return {
    columns: new Axis(MAX_COLUMN + 1, DEFAULT_COLUMN_WIDTH, cols),
    rows: new Axis(MAX_ROW + 1, DEFAULT_ROW_HEIGHT, rows),
    lastRow,
    lastColumn,
  };
}

/** The bottom-right corner the scroller must reach: cells, merges and laid-out lines. */
export function contentEnd(sheet: Worksheet, layout: SheetLayout): { bottom: number; right: number } {
  const cells = sheet.cellExtent();
  let bottom = Math.max(cells?.lastRow ?? 0, layout.lastRow);
  let right = Math.max(cells?.lastColumn ?? 0, layout.lastColumn);
  for (const m of sheet.merges) {
    bottom = Math.max(bottom, m.range.bottom);
    right = Math.max(right, m.range.right);
  }
  return { bottom, right };
}
