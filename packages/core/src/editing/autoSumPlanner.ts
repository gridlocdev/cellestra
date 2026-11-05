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
    if (!selection.isSingleCell && selectionHasNumbers(sheet, selection))
      return columnTotals(sheet, selection);
    const above = runAbove(sheet, activeRow, activeColumn);
    if (above) return [{ targetRow: activeRow, targetColumn: activeColumn, formula: sumFormula(above) }];
    const left = runLeft(sheet, activeRow, activeColumn);
    if (left) return [{ targetRow: activeRow, targetColumn: activeColumn, formula: sumFormula(left) }];
    return [];
  },
} as const;

function columnTotals(sheet: Worksheet, selection: CellRange): AutoSumProposal[] {
  const targetRow = selection.bottom + 1;
  if (targetRow >= SheetDimensions.maxRows) return [];
  const out: AutoSumProposal[] = [];
  for (let column = selection.left; column <= selection.right; column++) {
    let numeric = false;
    for (let row = selection.top; row <= selection.bottom && !numeric; row++) {
      numeric = sheet.getValue(row, column).kind === 'number';
    }
    if (!numeric) continue; // a label column inside the selection
    const range = CellRange.fromBounds(selection.top, column, selection.bottom, column);
    out.push({ targetRow, targetColumn: column, formula: sumFormula(range) });
  }
  return out;
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

function selectionHasNumbers(sheet: Worksheet, selection: CellRange): boolean {
  for (const [row, column] of selection.positions())
    if (sheet.getValue(row, column).kind === 'number') return true;
  return false;
}

const sumFormula = (range: CellRange): string => `=SUM(${range.toString()})`;
