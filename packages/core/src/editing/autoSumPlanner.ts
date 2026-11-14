import { CellRange } from '../model/cellRange.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
import type { Worksheet } from '../model/worksheet.ts';

/** A SUM to write: target cell and the formula text (with `=`). */
export interface AutoSumProposal {
  readonly targetRow: number;
  readonly targetColumn: number;
  readonly formula: string;
}

/**
 * AutoSum: a multi-cell selection holding numbers totals each numeric column into the row below;
 * otherwise the active cell sums the run of numbers directly above it, else directly to its left.
 */
export const AutoSumPlanner = {
  plan(sheet: Worksheet, selection: CellRange, activeRow: number, activeColumn: number): AutoSumProposal[] {
    if (!selection.isSingleCell) {
      const numeric = numericColumns(sheet, selection);
      if (numeric.size > 0) return columnTotals(selection, numeric);
    }
    const above = runAbove(sheet, activeRow, activeColumn);
    if (above) return [{ targetRow: activeRow, targetColumn: activeColumn, formula: sumFormula(above) }];
    const left = runLeft(sheet, activeRow, activeColumn);
    if (left) return [{ targetRow: activeRow, targetColumn: activeColumn, formula: sumFormula(left) }];
    return [];
  },
} as const;

function columnTotals(selection: CellRange, numeric: ReadonlySet<number>): AutoSumProposal[] {
  const targetRow = selection.bottom + 1;
  if (targetRow >= SheetDimensions.maxRows) return [];
  // Columns without a number are labels inside the selection.
  return [...numeric]
    .sort((a, b) => a - b)
    .map((column) => ({
      targetRow,
      targetColumn: column,
      formula: sumFormula(CellRange.fromBounds(selection.top, column, selection.bottom, column)),
    }));
}

function runAbove(sheet: Worksheet, row: number, column: number): CellRange | null {
  let top = row;
  while (top - 1 >= 0 && sheet.getValue(top - 1, column).kind === 'number') top--;
  return top === row ? null : CellRange.fromBounds(top, column, row - 1, column);
}

function runLeft(sheet: Worksheet, row: number, column: number): CellRange | null {
  let left = column;
  while (left - 1 >= 0 && sheet.getValue(row, left - 1).kind === 'number') left--;
  return left === column ? null : CellRange.fromBounds(row, left, row, column - 1);
}

/**
 * The columns of the selection that hold a number. Only the used part of the sheet is read, and a
 * selection bigger than the sheet's cell count is answered from the stored cells instead, so a
 * whole-sheet selection costs no more than the cells there are.
 */
function numericColumns(sheet: Worksheet, selection: CellRange): Set<number> {
  const out = new Set<number>();
  const extent = sheet.cellExtent();
  if (!extent) return out;
  const bottom = Math.min(selection.bottom, extent.lastRow);
  const right = Math.min(selection.right, extent.lastColumn);
  if (bottom < selection.top || right < selection.left) return out;
  const area = (bottom - selection.top + 1) * (right - selection.left + 1);
  if (area > sheet.cellCount) {
    for (const [row, column, cell] of sheet.enumerateCells())
      if (cell.value.kind === 'number' && selection.contains(row, column)) out.add(column);
    return out;
  }
  for (let column = selection.left; column <= right; column++) {
    for (let row = selection.top; row <= bottom; row++) {
      if (sheet.getValue(row, column).kind === 'number') {
        out.add(column);
        break;
      }
    }
  }
  return out;
}

const sumFormula = (range: CellRange): string => `=SUM(${range.toString()})`;
