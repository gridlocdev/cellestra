import { lineMoveCoord, lineMoveSpan, ReferenceAdjuster } from '../formulas/references/referenceAdjuster.ts';
import { CellRange } from '../model/cellRange.ts';
import { MergedRegion } from '../model/mergedRegion.ts';
import { cellKey, SheetDimensions } from '../model/sheetDimensions.ts';
import type { LineGroup, LineSpan, SheetOutline } from '../model/sheetOutline.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet, WorksheetLayout } from '../model/worksheet.ts';
import {
  EditRefusedError,
  EditResult,
  requireSheet,
  type ChangedCell,
  type EditCommand,
} from './editCommand.ts';
import { rewriteFormula } from './formulaRewrite.ts';
import type { StructuralAxis } from './structuralEditCommand.ts';
import {
  captureReferences,
  restoreReferences,
  transformReferences,
  transformTableKeys,
  type ReferenceSnapshot,
  type ReferenceTransform,
} from './workbookReferences.ts';

/**
 * Move whole rows (or columns) `first`…`last` one step up or down (left or right): they swap places
 * with the line next to them. Cells, line heights and widths, hidden lines and formatting go with
 * them, and every reference to their cells follows (formulas, names, validations, filters, charts).
 * A merge or table cut across the edge of the move refuses it.
 */
export class MoveLinesCommand implements EditCommand {
  readonly label: string;
  private readonly sheetName: string;
  private readonly axis: StructuralAxis;
  private readonly first: number;
  private readonly last: number;
  private readonly delta: -1 | 1;
  private layoutBefore: WorksheetLayout | null = null;
  private formulasBefore: [ChangedCell, string][] | null = null;
  private referencesBefore: ReferenceSnapshot | null = null;

  constructor(sheetName: string, axis: StructuralAxis, first: number, last: number, delta: -1 | 1) {
    this.sheetName = sheetName;
    this.axis = axis;
    this.first = first;
    this.last = last;
    this.delta = delta;
    this.label = `Move ${axis === 'row' ? 'Rows' : 'Columns'}`;
  }

  /** Why the move can't happen on this workbook, or null when it can. */
  refusal(workbook: Workbook): string | null {
    const sheet = workbook.findSheet(this.sheetName);
    if (!sheet) return null;
    const rows = this.axis === 'row';
    const limit = rows ? SheetDimensions.maxRows : SheetDimensions.maxColumns;
    const lines = rows ? 'rows' : 'columns';
    if (this.first + this.delta < 0 || this.last + this.delta >= limit)
      return `Can't move the ${lines} past the edge of the sheet.`;
    const split = (range: CellRange) => this.span(range, rows) === 'split';
    if (sheet.merges.some((m) => split(m.range)))
      return `Can't move these ${lines}: it would split merged cells. Unmerge them first.`;
    if (sheet.tables.some((t) => split(t.range)))
      return `Can't move these ${lines} across the edge of a table.`;
    return null;
  }

  apply(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
    const refusal = this.refusal(workbook);
    if (refusal) throw new EditRefusedError(refusal);
    const layout = sheet.captureLayout();
    this.layoutBefore = layout;
    this.referencesBefore = captureReferences(workbook);
    this.moveCells(sheet, (line) => this.map(line));
    sheet.restoreLayout(this.movedLayout(layout));
    this.formulasBefore = this.adjustFormulas(workbook);
    transformReferences(workbook, this.transform);
    return EditResult.structuralChange;
  }

  undo(workbook: Workbook): EditResult {
    const { layoutBefore, formulasBefore, referencesBefore } = this;
    if (!layoutBefore || !formulasBefore || !referencesBefore) return EditResult.none;
    const sheet = requireSheet(workbook, this.sheetName);
    // Formulas first, while the cells still sit where the move put them.
    for (const [ref, formula] of formulasBefore) {
      const cell = workbook.findSheet(ref.sheet)?.find(ref.row, ref.column);
      if (cell) cell.formula = formula;
    }
    this.moveCells(sheet, (line) => this.unmap(line));
    sheet.restoreLayout(layoutBefore);
    restoreReferences(workbook, referencesBefore);
    return EditResult.structuralChange;
  }

  private moveCells(sheet: Worksheet, map: (line: number) => number): void {
    const rows = this.axis === 'row';
    sheet.moveCells((row, column) => (rows ? cellKey(map(row), column) : cellKey(row, map(column))));
  }

  private map(line: number): number {
    return lineMoveCoord(line, this.first, this.last, this.delta);
  }

  /** The inverse: the moved lines step back, and the swapped line returns to its place. */
  private unmap(line: number): number {
    const first = this.first + this.delta;
    const last = this.last + this.delta;
    return lineMoveCoord(line, first, last, this.delta === 1 ? -1 : 1);
  }

  private span(range: CellRange, rows: boolean): 'moves' | 'stays' | 'split' {
    return lineMoveSpan(
      rows ? range.top : range.left,
      rows ? range.bottom : range.right,
      this.first,
      this.last,
      this.delta,
    );
  }

  private movedRange(range: CellRange, rows: boolean): CellRange {
    if (this.span(range, rows) !== 'moves') return range;
    return rows
      ? CellRange.fromBounds(this.map(range.top), range.left, this.map(range.bottom), range.right)
      : CellRange.fromBounds(range.top, this.map(range.left), range.bottom, this.map(range.right));
  }

  private movedLayout(before: WorksheetLayout): WorksheetLayout {
    const rows = this.axis === 'row';
    const remap = <T>(source: ReadonlyMap<number, T>): Map<number, T> =>
      new Map([...source].map(([k, v]) => [this.map(k), v]));
    return {
      rows: rows ? remap(before.rows) : new Map(before.rows),
      columns: rows ? new Map(before.columns) : remap(before.columns),
      merges: before.merges.map((m) => new MergedRegion(this.movedRange(m.range, rows))),
      tables: before.tables.map((t) =>
        transformTableKeys({ ...t, range: this.movedRange(t.range, rows) }, this.sheetName, this.transform),
      ),
      frozenRows: before.frozenRows,
      frozenColumns: before.frozenColumns,
      outline: this.movedOutline(before.outline, rows),
    };
  }

  /** Groups and print titles along the axis move only when they sit wholly in the moved lines. */
  private movedOutline(outline: SheetOutline, rows: boolean): SheetOutline {
    const move = <T extends LineSpan>(span: T): T =>
      lineMoveSpan(span.first, span.last, this.first, this.last, this.delta) === 'moves'
        ? { ...span, first: this.map(span.first), last: this.map(span.last) }
        : span;
    const groups = (list: readonly LineGroup[]): LineGroup[] => list.map(move);
    return rows
      ? {
          ...outline,
          rowGroups: groups(outline.rowGroups),
          headerRows: outline.headerRows && move(outline.headerRows),
        }
      : {
          ...outline,
          columnGroups: groups(outline.columnGroups),
          headerColumns: outline.headerColumns && move(outline.headerColumns),
        };
  }

  private readonly transform: ReferenceTransform = (expr, home) =>
    ReferenceAdjuster.forLineMove(
      expr,
      home,
      this.sheetName,
      this.axis === 'row',
      this.first,
      this.last,
      this.delta,
    );

  /** Rewrite every formula in the workbook; returns the ones that changed, with their old text. */
  private adjustFormulas(workbook: Workbook): [ChangedCell, string][] {
    const changed: [ChangedCell, string][] = [];
    for (const sheet of workbook.worksheets) {
      const home = sheet.name;
      for (const [row, column, cell] of sheet.enumerateCells()) {
        if (cell.formula === null) continue;
        const rewritten = rewriteFormula(cell.formula, (expr) => this.transform(expr, home));
        if (rewritten === cell.formula) continue;
        changed.push([{ sheet: home, row, column }, cell.formula]);
        cell.formula = rewritten;
      }
    }
    return changed;
  }
}
