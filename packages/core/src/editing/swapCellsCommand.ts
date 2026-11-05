import { ReferenceAdjuster } from '../formulas/references/referenceAdjuster.ts';
import type { Cell } from '../model/cell.ts';
import type { CellRange } from '../model/cellRange.ts';
import type { Workbook } from '../model/workbook.ts';
import { CellSnapshots } from './cellSnapshots.ts';
import { EditResult, requireSheet, type ChangedCell, type EditCommand } from './editCommand.ts';
import { rewriteFormula } from './formulaRewrite.ts';

/**
 * Exchange block A with the same-sized block offset by (rowDelta, columnDelta) — a drag-move onto
 * occupied cells. Formulas referencing either block are re-pointed so they follow the cells. With
 * `affectedFormulas` (from the dependency graph) only those are rewritten, before the swap; otherwise
 * every formula in the workbook is scanned after it. The blocks must not overlap.
 */
export class SwapCellsCommand implements EditCommand {
  readonly label = 'Move';
  private readonly sheetName: string;
  private readonly blockA: CellRange;
  private readonly rowDelta: number;
  private readonly columnDelta: number;
  private readonly affectedFormulas: readonly ChangedCell[] | null;
  private before: CellSnapshots | null = null;

  constructor(
    sheetName: string,
    blockA: CellRange,
    rowDelta: number,
    columnDelta: number,
    affectedFormulas: readonly ChangedCell[] | null = null,
  ) {
    this.sheetName = sheetName;
    this.blockA = blockA;
    this.rowDelta = rowDelta;
    this.columnDelta = columnDelta;
    this.affectedFormulas = affectedFormulas;
  }

  apply(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
    const before = (this.before = new CellSnapshots());
    const { top, left, rowCount: rows, columnCount: cols } = this.blockA;
    const dr = this.rowDelta;
    const dc = this.columnDelta;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        before.capture(sheet, top + r, left + c);
        before.capture(sheet, top + r + dr, left + c + dc);
      }
    }
    // Rewritten in place before the swap, so a referencing cell inside a block carries its new formula along.
    if (this.affectedFormulas) this.rewriteReferences(workbook, this.affectedFormulas);

    const a: (Cell | null)[] = [];
    const b: (Cell | null)[] = [];
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        a.push(sheet.find(top + r, left + c)?.clone() ?? null);
        b.push(sheet.find(top + r + dr, left + c + dc)?.clone() ?? null);
      }
    }
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        sheet.putCell(top + r, left + c, b[r * cols + c]);
        sheet.putCell(top + r + dr, left + c + dc, a[r * cols + c]);
      }
    }
    if (!this.affectedFormulas) this.rewriteReferences(workbook, allFormulaCells(workbook));
    return EditResult.cells(before.refs());
  }

  undo(workbook: Workbook): EditResult {
    if (!this.before) return EditResult.none;
    this.before.restore(workbook);
    return EditResult.cells(this.before.refs());
  }

  private rewriteReferences(workbook: Workbook, cells: readonly ChangedCell[]): void {
    for (const ref of cells) {
      const sheet = workbook.findSheet(ref.sheet);
      const cell = sheet?.find(ref.row, ref.column);
      if (!sheet || !cell || cell.formula === null) continue;
      const rewritten = rewriteFormula(cell.formula, (expr) =>
        ReferenceAdjuster.forSwap(
          expr,
          ref.sheet,
          this.sheetName,
          this.blockA,
          this.rowDelta,
          this.columnDelta,
        ),
      );
      if (rewritten !== cell.formula) {
        this.before!.capture(sheet, ref.row, ref.column);
        cell.formula = rewritten;
      }
    }
  }
}

function allFormulaCells(workbook: Workbook): ChangedCell[] {
  const out: ChangedCell[] = [];
  for (const sheet of workbook.worksheets) {
    for (const [row, column, cell] of sheet.enumerateCells()) {
      if (cell.formula !== null) out.push({ sheet: sheet.name, row, column });
    }
  }
  return out;
}
