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
    do {
      r2 += rowDelta;
      c2 += columnDelta;
    } while (!atBoundary(r2, c2) && !filled(r2, c2));
    if (filled(r2, c2)) return [r2, c2];

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

/** True when target lies strictly past origin on the way to boundary. */
function isBetween(origin: number, target: number, boundary: number): boolean {
  return boundary >= origin ? target > origin && target <= boundary : target < origin && target >= boundary;
}
