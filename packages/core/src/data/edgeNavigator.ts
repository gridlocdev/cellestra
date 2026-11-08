import { SheetDimensions } from '../model/sheetDimensions.ts';
import type { Worksheet } from '../model/worksheet.ts';

/** Excel-style ⌘+Arrow "jump to edge" (PRD parity). */
export const EdgeNavigator = {
  /**
   * Where ⌘+Arrow from (row, column) lands moving by (rowDelta, columnDelta) — one of them ±1. Inside
   * a run: its last filled cell. Facing a gap or from an empty cell: the next filled cell, else the
   * used-range edge when it is ahead, else the sheet edge.
   */
  findEdge(
    sheet: Worksheet,
    row: number,
    column: number,
    rowDelta: number,
    columnDelta: number,
  ): [number, number] {
    const boundary =
      rowDelta !== 0
        ? rowDelta > 0
          ? SheetDimensions.maxRows - 1
          : 0
        : columnDelta > 0
          ? SheetDimensions.maxColumns - 1
          : 0;
    const filled = (r: number, c: number): boolean => sheet.find(r, c) !== undefined;
    const axis = (r: number, c: number): number => (rowDelta !== 0 ? r : c);
    const atBoundary = (r: number, c: number): boolean => axis(r, c) === boundary;

    if (atBoundary(row, column)) return [row, column];
    let r2 = row;
    let c2 = column;
    if (filled(row, column) && filled(row + rowDelta, column + columnDelta)) {
      while (!atBoundary(r2, c2) && filled(r2 + rowDelta, c2 + columnDelta)) {
        r2 += rowDelta;
        c2 += columnDelta;
      }
      return [r2, c2];
    }
    const next = nextFilled(sheet, row, column, rowDelta, columnDelta, boundary);
    if (next) return next;
    if (rowDelta !== 0) r2 = boundary;
    else c2 = boundary;

    // Ran off the data: prefer the used-range edge ahead of the start, not row 1,048,576.
    const used = sheet.getUsedRange();
    if (used) {
      const edge =
        rowDelta > 0
          ? Math.max(row, used.bottom)
          : rowDelta < 0
            ? Math.min(row, used.top)
            : columnDelta > 0
              ? Math.max(column, used.right)
              : Math.min(column, used.left);
      if (axis(r2, c2) !== edge && isBetween(axis(row, column), edge, boundary)) {
        return rowDelta !== 0 ? [edge, column] : [row, edge];
      }
    }
    return [r2, c2];
  },
} as const;

/**
 * The first stored cell past (row, column) in the direction of travel, up to `boundary`, else null.
 * Nothing lies past the sheet's last stored cell, and when the walk would still be longer than the
 * number of stored cells, the cells are searched instead: an empty column costs no million steps.
 */
function nextFilled(
  sheet: Worksheet,
  row: number,
  column: number,
  rowDelta: number,
  columnDelta: number,
  boundary: number,
): [number, number] | null {
  const extent = sheet.cellExtent();
  if (!extent) return null;
  const vertical = rowDelta !== 0;
  const step = vertical ? rowDelta : columnDelta;
  const from = vertical ? row : column;
  const stop = step > 0 ? Math.min(boundary, vertical ? extent.lastRow : extent.lastColumn) : boundary;
  const at = (i: number): [number, number] => (vertical ? [i, column] : [row, i]);
  if (Math.abs(stop - from) <= sheet.cellCount) {
    for (let i = from + step; step > 0 ? i <= stop : i >= stop; i += step) {
      const [r, c] = at(i);
      if (sheet.find(r, c) !== undefined) return [r, c];
    }
    return null;
  }
  let best: number | null = null;
  for (const [r, c] of sheet.enumerateCells()) {
    if (vertical ? c !== column : r !== row) continue;
    const i = vertical ? r : c;
    if ((i - from) * step <= 0 || (i - stop) * step > 0) continue;
    if (best === null || (i - best) * step < 0) best = i;
  }
  return best === null ? null : at(best);
}

/** True when target lies strictly past origin on the way to boundary. */
function isBetween(origin: number, target: number, boundary: number): boolean {
  return boundary >= origin ? target > origin && target <= boundary : target < origin && target >= boundary;
}
