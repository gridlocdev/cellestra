import type { Expr } from '../formulas/ast/expr.ts';
import { ReferenceAdjuster } from '../formulas/references/referenceAdjuster.ts';
import type { Workbook } from '../model/workbook.ts';
import { captureFormulas, restoreFormulas } from './deleteSheetCommand.ts';
import { EditResult, requireSheet, type ChangedCell, type EditCommand } from './editCommand.ts';
import { rewriteFormula } from './formulaRewrite.ts';
import {
  captureReferences,
  restoreReferences,
  transformReferences,
  type ReferenceSnapshot,
} from './workbookReferences.ts';

const same = (a: string | null | undefined, b: string): boolean =>
  a !== null && a !== undefined && a.toLowerCase() === b.toLowerCase();

/**
 * Rename a sheet and every reference to it: formulas that name it (`Old.A1` → `New.A1`), named-range
 * targets and sheet scopes, validation list sources, and chart ranges. Throws (from
 * `Workbook.renameSheet`) on an invalid or duplicate name, before anything changes. Undo renames it
 * back and restores the captured formulas and references.
 */
export class RenameSheetCommand implements EditCommand {
  readonly label = 'Rename Sheet';
  private readonly oldName: string;
  private readonly newName: string;
  private formulasBefore: [ChangedCell, string][] | null = null;
  private referencesBefore: ReferenceSnapshot | null = null;

  constructor(oldName: string, newName: string) {
    this.oldName = oldName;
    this.newName = newName;
  }

  apply(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.oldName);
    const formulas = captureFormulas(workbook, null);
    const references = captureReferences(workbook);
    workbook.renameSheet(sheet, this.newName);
    this.formulasBefore = formulas;
    this.referencesBefore = references;

    const { oldName, newName } = this;
    const transform = (expr: Expr): Expr => ReferenceAdjuster.forRenameSheet(expr, oldName, newName);
    for (const s of workbook.worksheets) {
      for (const [, , cell] of s.enumerateCells()) {
        if (cell.formula !== null) cell.formula = rewriteFormula(cell.formula, transform);
      }
    }
    transformReferences(workbook, transform);
    workbook.replaceNamedRanges(
      workbook.namedRanges.map((n) => (same(n.scope, oldName) ? { ...n, scope: newName } : n)),
    );
    return EditResult.structuralChange;
  }

  undo(workbook: Workbook): EditResult {
    if (!this.formulasBefore || !this.referencesBefore) return EditResult.none;
    // Straight back to the old name: it was valid before, even if loaded from a file our rules reject.
    requireSheet(workbook, this.newName).name = this.oldName;
    // The snapshots were taken under the old name, which the sheet carries again.
    restoreFormulas(workbook, this.formulasBefore);
    restoreReferences(workbook, this.referencesBefore);
    return EditResult.structuralChange;
  }
}
