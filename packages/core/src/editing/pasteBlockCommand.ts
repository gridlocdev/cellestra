import { ReferenceAdjuster } from '../formulas/references/referenceAdjuster.ts';
import { Cell } from '../model/cell.ts';
import { CellRange } from '../model/cellRange.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { StylePool } from '../styling/stylePool.ts';
import { CellValue } from '../values/cellValue.ts';
import type { ClipboardBlock, ClipboardCell } from './clipboardBlock.ts';
import { CellSnapshots } from './cellSnapshots.ts';
import { EditRefusedError, EditResult, requireSheet, type EditCommand } from './editCommand.ts';
import { rewriteFormula } from './formulaRewrite.ts';
import { PastePlanner, type PastePlan } from './pastePlanner.ts';

const sameSheet = (a: string, b: string): boolean => a.toLowerCase() === b.toLowerCase();

/** What a paste carries: everything, the computed values, only formats, or formulas without formats. */
export type PasteMode = 'all' | 'values' | 'formats' | 'formulas';

/** A value a paste or fill is about to write — the input to validation pre-flight. */
export interface ProjectedWrite {
  readonly row: number;
  readonly column: number;
  readonly value: CellValue;
  readonly isFormula: boolean;
}

/**
 * Paste (or move, for a cut) a clipboard block onto a destination selection, planned by
 * `PastePlanner`. A copy shifts relative formula references by each anchor's offset; a move writes
 * formulas verbatim and then re-points every formula in the workbook that referenced the old cells.
 */
export class PasteBlockCommand implements EditCommand {
  readonly label: string;
  private readonly destSheet: string;
  private readonly block: ClipboardBlock;
  private readonly plan: PastePlan;
  private readonly mode: PasteMode;
  private before: CellSnapshots | null = null;

  /** `destination` is the selection, or a (row, column) top-left to expand from. */
  constructor(
    destSheet: string,
    destination: CellRange | readonly [row: number, column: number],
    block: ClipboardBlock,
    mode: PasteMode = 'all',
  ) {
    this.destSheet = destSheet;
    this.block = block;
    this.mode = mode;
    const range =
      destination instanceof CellRange
        ? destination
        : CellRange.fromBounds(destination[0], destination[1], destination[0], destination[1]);
    this.plan = PastePlanner.plan(block.rows, block.columns, range, block.isCut);
    this.label = block.isCut ? 'Move' : mode === 'all' ? 'Paste' : 'Paste Special';
  }

  get shapeMismatch(): boolean {
    return this.plan.shapeMismatch;
  }

  apply(workbook: Workbook): EditResult {
    this.requireFit();
    const dest = requireSheet(workbook, this.destSheet);
    const source = workbook.findSheet(this.block.sourceSheet);
    const b = this.block;
    this.before = new CellSnapshots();
    if (b.isCut && source)
      for (const [row, column] of b.sourceRange.positions()) this.snapshot(source, row, column);
    for (const [ar, ac] of this.plan.anchors) {
      for (let r = 0; r < b.rows; r++)
        for (let c = 0; c < b.columns; c++) this.snapshot(dest, ar + r, ac + c);
    }
    if (b.isCut && source)
      for (const [row, column] of b.sourceRange.positions()) source.putCell(row, column, null);

    for (const [ar, ac] of this.plan.anchors) {
      const rowDelta = ar - b.sourceRange.top;
      const columnDelta = ac - b.sourceRange.left;
      for (let r = 0; r < b.rows; r++) {
        for (let c = 0; c < b.columns; c++) {
          this.writeCell(dest, ar + r, ac + c, b.cells[r]![c] ?? null, rowDelta, columnDelta);
        }
      }
    }

    // A move relocates the cells, so every formula that pointed at the old location follows them —
    // the destination cells included (a block where B2 = B1 keeps pointing inside itself).
    const first = this.plan.anchors[0];
    if (b.isCut && first)
      this.adjustMovedReferences(workbook, first[0] - b.sourceRange.top, first[1] - b.sourceRange.left);
    return EditResult.cells(this.before.refs());
  }

  undo(workbook: Workbook): EditResult {
    if (!this.before) return EditResult.none;
    this.before.restore(workbook);
    return EditResult.cells(this.before.refs());
  }

  /** The values this paste would write (none for a formats-only paste), without applying it. */
  *projectWrites(): IterableIterator<ProjectedWrite> {
    if (this.mode === 'formats' || !this.fits) return;
    const b = this.block;
    for (const [ar, ac] of this.plan.anchors) {
      for (let r = 0; r < b.rows; r++) {
        for (let c = 0; c < b.columns; c++) {
          const source = b.cells[r]![c];
          if (!source) continue;
          yield {
            row: ar + r,
            column: ac + c,
            value: source.value,
            isFormula: this.mode !== 'values' && source.formula !== null,
          };
        }
      }
    }
  }

  /** False when the paste would run past the sheet's last row or column (`apply` refuses it). */
  get fits(): boolean {
    const { rows, columns } = this.block;
    return this.plan.anchors.every(([row, column]) => PasteBlockCommand.fits(row, column, rows, columns));
  }

  /** Whether a `rows` × `columns` block placed at (row, column) stays inside the sheet. */
  static fits(row: number, column: number, rows: number, columns: number): boolean {
    return row + rows <= SheetDimensions.maxRows && column + columns <= SheetDimensions.maxColumns;
  }

  static readonly doesNotFit = "The clipboard doesn't fit here: it would run past the edge of the sheet.";

  /** Refuse a paste that would run past the sheet's edge (the cells would wrap otherwise). */
  private requireFit(): void {
    if (!this.fits) throw new EditRefusedError(PasteBlockCommand.doesNotFit);
  }

  private adjustMovedReferences(workbook: Workbook, rowDelta: number, columnDelta: number): void {
    const b = this.block;
    for (const sheet of workbook.worksheets) {
      for (const [row, column, cell] of [...sheet.enumerateCells()]) {
        if (cell.formula === null) continue;
        const rewritten = rewriteFormula(cell.formula, (expr) =>
          ReferenceAdjuster.forMove(
            expr,
            sheet.name,
            b.sourceSheet,
            b.sourceRange,
            this.destSheet,
            rowDelta,
            columnDelta,
          ),
        );
        if (rewritten !== cell.formula) {
          this.snapshot(sheet, row, column);
          cell.formula = rewritten;
        }
      }
    }
  }

  private snapshot(sheet: Worksheet, row: number, column: number): void {
    this.before!.capture(sheet, row, column);
  }

  private writeCell(
    dest: Worksheet,
    row: number,
    column: number,
    source: ClipboardCell | null,
    rowDelta: number,
    columnDelta: number,
  ): void {
    // Formats keeps the destination's content and only re-skins it; a blank source resets to default.
    if (this.mode === 'formats') {
      const cell = dest.find(row, column)?.clone() ?? new Cell();
      cell.styleId = source?.styleId ?? StylePool.defaultId;
      dest.putCell(row, column, cell);
      return;
    }
    if (!source) {
      // All wipes the cell; Values/Formulas clear only the content and keep the destination's style.
      if (this.mode === 'all') {
        dest.putCell(row, column, null);
      } else {
        const existing = dest.find(row, column);
        if (existing) {
          existing.formula = null;
          existing.value = CellValue.empty;
          dest.putCell(row, column, existing);
        }
      }
      return;
    }
    let formula = source.formula;
    if (formula !== null && !this.block.isCut) {
      formula = rewriteFormula(formula, (expr) => ReferenceAdjuster.forCopy(expr, columnDelta, rowDelta));
    } else if (formula !== null && !sameSheet(this.block.sourceSheet, this.destSheet)) {
      // A cut moved to another sheet: pin its unqualified references to the source sheet first, so
      // they keep meaning the source's cells; the move fix-up then re-homes the ones inside the block.
      const sourceSheet = this.block.sourceSheet;
      formula = rewriteFormula(formula, (expr) => ReferenceAdjuster.qualify(expr, sourceSheet));
    }
    const keepDestStyle = this.mode !== 'all';
    const styleId = keepDestStyle ? (dest.find(row, column)?.styleId ?? StylePool.defaultId) : source.styleId;
    const cell = new Cell();
    cell.value = source.value;
    cell.formula = this.mode === 'values' ? null : formula;
    cell.formulaSource = source.formulaSource;
    cell.styleId = styleId;
    dest.putCell(row, column, cell);
  }
}
