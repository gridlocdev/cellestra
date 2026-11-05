import { ReferenceAdjuster } from '../formulas/references/referenceAdjuster.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { EditResult, requireSheet, type ChangedCell, type EditCommand } from './editCommand.ts';
import { rewriteFormula } from './formulaRewrite.ts';

/**
 * Delete a sheet; references to it elsewhere become `#REF!`. Undo re-inserts the same sheet object
 * at its old index and restores every other sheet's formula text. Throws on the last sheet.
 */
export class DeleteSheetCommand implements EditCommand {
  readonly label = 'Delete Sheet';
  private readonly sheetName: string;
  private index = 0;
  private removed: Worksheet | null = null;
  private otherFormulasBefore: [ChangedCell, string][] | null = null;

  constructor(sheetName: string) {
    this.sheetName = sheetName;
  }

  apply(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
    this.index = workbook.indexOf(sheet);
    this.removed = sheet;
    this.otherFormulasBefore = captureFormulas(workbook, sheet);
    workbook.removeSheet(sheet);
    for (const other of workbook.worksheets) {
      for (const [, , cell] of other.enumerateCells()) {
        if (cell.formula === null) continue;
        cell.formula = rewriteFormula(cell.formula, (expr) =>
          ReferenceAdjuster.forDeleteSheet(expr, this.sheetName, other.name),
        );
      }
    }
    return EditResult.structuralChange;
  }

  undo(workbook: Workbook): EditResult {
    if (!this.removed || !this.otherFormulasBefore) return EditResult.none;
    workbook.insertSheet(this.index, this.removed);
    restoreFormulas(workbook, this.otherFormulasBefore);
    return EditResult.structuralChange;
  }
}

/** Every formula's text on every sheet except `except`, for structural undo. */
export function captureFormulas(workbook: Workbook, except: Worksheet | null): [ChangedCell, string][] {
  const out: [ChangedCell, string][] = [];
  for (const sheet of workbook.worksheets) {
    if (sheet === except) continue;
    for (const [row, column, cell] of sheet.enumerateCells()) {
      if (cell.formula !== null) out.push([{ sheet: sheet.name, row, column }, cell.formula]);
    }
  }
  return out;
}

export function restoreFormulas(
  workbook: Workbook,
  formulas: readonly (readonly [ChangedCell, string])[],
): void {
  for (const [ref, formula] of formulas) {
    const cell = workbook.findSheet(ref.sheet)?.find(ref.row, ref.column);
    if (cell) cell.formula = formula;
  }
}
