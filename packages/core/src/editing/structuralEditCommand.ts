import { ReferenceAdjuster } from '../formulas/references/referenceAdjuster.ts';
import type { Table } from '../data/types.ts';
import { Cell } from '../model/cell.ts';
import { CellRange } from '../model/cellRange.ts';
import { Column, Row } from '../model/lines.ts';
import { MergedRegion } from '../model/mergedRegion.ts';
import { cellKey, keyColumn, keyRow, SheetDimensions } from '../model/sheetDimensions.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet, WorksheetState } from '../model/worksheet.ts';
import { captureFormulas, restoreFormulas } from './deleteSheetCommand.ts';
import {
  EditRefusedError,
  EditResult,
  requireSheet,
  type ChangedCell,
  type EditCommand,
} from './editCommand.ts';
import { rewriteFormula } from './formulaRewrite.ts';
import {
  captureReferences,
  restoreReferences,
  transformReferences,
  transformTableKeys,
  type ReferenceSnapshot,
  type ReferenceTransform,
} from './workbookReferences.ts';

export type StructuralAxis = 'row' | 'column';
export type StructuralOperation = 'insert' | 'delete';

/**
 * Insert or delete whole rows/columns: cells, layout, merges, and tables shift; inserted lines
 * inherit the formatting of the line before them (PRD §8.14); every formula in the workbook is
 * adjusted (deleted references become `#REF!`), and so are named ranges, data validations, the
 * AutoFilter and sort keys, and chart ranges. Undo restores the captured state.
 */
export class StructuralEditCommand implements EditCommand {
  readonly label: string;
  private readonly sheetName: string;
  private readonly axis: StructuralAxis;
  private readonly operation: StructuralOperation;
  private readonly at: number;
  private readonly count: number;
  private editedBefore: WorksheetState | null = null;
  private otherFormulasBefore: [ChangedCell, string][] | null = null;
  private referencesBefore: ReferenceSnapshot | null = null;

  constructor(
    sheetName: string,
    axis: StructuralAxis,
    operation: StructuralOperation,
    at: number,
    count: number,
  ) {
    this.sheetName = sheetName;
    this.axis = axis;
    this.operation = operation;
    this.at = at;
    this.count = count;
    this.label = `${operation === 'insert' ? 'Insert' : 'Delete'} ${axis === 'row' ? 'Rows' : 'Columns'}`;
  }

  apply(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
    this.refuseDataLoss(sheet);
    this.editedBefore = sheet.capture();
    this.otherFormulasBefore = captureFormulas(workbook, sheet);
    this.referencesBefore = captureReferences(workbook);
    sheet.restore(this.buildShiftedState(this.editedBefore));
    this.adjustFormulas(workbook);
    transformReferences(workbook, this.transform);
    return EditResult.structuralChange;
  }

  undo(workbook: Workbook): EditResult {
    if (!this.editedBefore || !this.otherFormulasBefore || !this.referencesBefore) return EditResult.none;
    requireSheet(workbook, this.sheetName).restore(this.editedBefore);
    restoreFormulas(workbook, this.otherFormulasBefore);
    restoreReferences(workbook, this.referencesBefore);
    return EditResult.structuralChange;
  }

  /**
   * An insert pushes the last lines off the sheet. Refuse when that would lose a value, a formula or
   * a comment, or cut a merge or table short, rather than drop them silently (formatting alone may go).
   */
  private refuseDataLoss(sheet: Worksheet): void {
    if (this.operation !== 'insert') return;
    const rows = this.axis === 'row';
    const firstLost = Math.max(this.at, this.limit - this.count);
    const extent = sheet.cellExtent();
    if (extent && (rows ? extent.lastRow : extent.lastColumn) >= firstLost) {
      for (const [row, column, cell] of sheet.enumerateCells()) {
        if ((rows ? row : column) < firstLost) continue;
        if (cell.formula !== null || cell.value.kind !== 'empty' || cell.unmodeled !== null)
          throw new EditRefusedError(this.lossMessage);
      }
    }
    const reaches = (range: CellRange) => (rows ? range.bottom : range.right) >= firstLost;
    if (sheet.merges.some((m) => reaches(m.range)) || sheet.tables.some((t) => reaches(t.range)))
      throw new EditRefusedError(this.lossMessage);
  }

  private get lossMessage(): string {
    const lines = this.axis === 'row' ? 'rows' : 'columns';
    return `Can't insert ${lines}: it would push data off the end of the sheet. Delete or move the data in the last ${lines} first.`;
  }

  // ---- The post-shift sheet state -------------------------------------------

  private buildShiftedState(before: WorksheetState): WorksheetState {
    const rows = this.axis === 'row';
    const cells = new Map<number, Cell>();
    for (const [key, cell] of before.cells) {
      const row = keyRow(key);
      const column = keyColumn(key);
      const mapped = this.mapCoord(rows ? row : column);
      if (mapped === null) continue;
      cells.set(rows ? cellKey(mapped, column) : cellKey(row, mapped), cell);
    }
    const shiftedRows = rows ? this.shiftMap(before.rows) : new Map(before.rows);
    const shiftedColumns = rows ? new Map(before.columns) : this.shiftMap(before.columns);
    // Inserting at the very top/left has no source line, so nothing to inherit.
    if (this.operation === 'insert' && this.at > 0) {
      this.inheritFormatting(before, cells, shiftedRows, shiftedColumns, rows);
    }
    return {
      cells,
      rows: shiftedRows,
      columns: shiftedColumns,
      merges: this.shiftMerges(before.merges, rows),
      tables: this.shiftTables(before.tables, rows),
      frozenRows: before.frozenRows,
      frozenColumns: before.frozenColumns,
    };
  }

  /**
   * Tables overlapping the edit resize along its axis; one wholly inside a deleted span goes. Their
   * filter and sort keys (absolute columns) follow their columns. Deleting a table's header row (or
   * header column) leaves it without one, rather than promoting its first data row.
   */
  private shiftTables(source: readonly Table[], rows: boolean): Table[] {
    const out: Table[] = [];
    for (const t of source) {
      const range = this.shiftRange(t.range, rows);
      if (!range) continue;
      let shifted: Table = { ...t, range };
      if (this.deletes(rows ? t.range.top : t.range.left)) {
        if (rows && t.hasHeader) {
          shifted = {
            ...shifted,
            hasHeader: false,
            filter: t.filter ? { ...t.filter, hasHeader: false } : t.filter,
            sort: t.sort ? { ...t.sort, hasHeader: false } : t.sort,
          };
        }
        if (!rows && t.hasHeaderColumn) shifted = { ...shifted, hasHeaderColumn: false };
      }
      out.push(transformTableKeys(shifted, this.sheetName, this.transform));
    }
    return out;
  }

  private deletes(line: number): boolean {
    return this.operation === 'delete' && line >= this.at && line < this.at + this.count;
  }

  private shiftMerges(source: readonly MergedRegion[], rows: boolean): MergedRegion[] {
    const out: MergedRegion[] = [];
    for (const m of source) {
      const range = this.shiftRange(m.range, rows);
      if (range && range.cellCount > 1) out.push(new MergedRegion(range));
    }
    return out;
  }

  private shiftRange(range: CellRange, rows: boolean): CellRange | null {
    const surviving = this.survivingExtremes(
      rows ? range.top : range.left,
      rows ? range.bottom : range.right,
    );
    const span = surviving && this.mapSpan(surviving[0], surviving[1]);
    if (!span) return null; // entirely deleted or pushed off the sheet
    const [lo, hi] = span;
    return rows
      ? CellRange.fromBounds(lo, range.left, hi, range.right)
      : CellRange.fromBounds(range.top, lo, range.bottom, hi);
  }

  private inheritFormatting(
    before: WorksheetState,
    cells: Map<number, Cell>,
    shiftedRows: Map<number, Row>,
    shiftedColumns: Map<number, Column>,
    rows: boolean,
  ): void {
    const source = this.at - 1;
    const end = this.at + this.count;
    if (rows) {
      const src = before.rows.get(source);
      if (src && src.styleId !== 0) {
        for (let r = this.at; r < end; r++)
          shiftedRows.set(r, Object.assign(new Row(), { styleId: src.styleId }));
      }
    } else {
      const src = before.columns.get(source);
      if (src && src.styleId !== 0) {
        for (let c = this.at; c < end; c++)
          shiftedColumns.set(c, Object.assign(new Column(), { styleId: src.styleId }));
      }
    }
    // The per-cell styles along the source line, as empty styled cells.
    for (const [key, cell] of before.cells) {
      const row = keyRow(key);
      const column = keyColumn(key);
      if ((rows ? row : column) !== source || cell.styleId === 0) continue;
      const cross = rows ? column : row;
      for (let i = this.at; i < end; i++) {
        const k = rows ? cellKey(i, cross) : cellKey(cross, i);
        if (!cells.has(k)) {
          const styled = new Cell();
          styled.styleId = cell.styleId;
          cells.set(k, styled);
        }
      }
    }
  }

  private shiftMap<T>(source: ReadonlyMap<number, T>): Map<number, T> {
    const out = new Map<number, T>();
    for (const [key, value] of source) {
      const mapped = this.mapCoord(key);
      if (mapped !== null) out.set(mapped, value);
    }
    return out;
  }

  // ---- Formulas -------------------------------------------------------------

  /** This edit's reference fix-up, as seen from a formula (or rule) on `home`. */
  private readonly transform: ReferenceTransform = (expr, home) => {
    const { at, count, sheetName } = this;
    if (this.axis === 'row') {
      return this.operation === 'insert'
        ? ReferenceAdjuster.forInsertRows(expr, home, sheetName, at, count)
        : ReferenceAdjuster.forDeleteRows(expr, home, sheetName, at, count);
    }
    return this.operation === 'insert'
      ? ReferenceAdjuster.forInsertColumns(expr, home, sheetName, at, count)
      : ReferenceAdjuster.forDeleteColumns(expr, home, sheetName, at, count);
  };

  private adjustFormulas(workbook: Workbook): void {
    for (const sheet of workbook.worksheets) {
      const home = sheet.name;
      for (const [, , cell] of sheet.enumerateCells()) {
        if (cell.formula === null) continue;
        cell.formula = rewriteFormula(cell.formula, (expr) => this.transform(expr, home));
      }
    }
  }

  // ---- Coordinate mapping (mirrors ReferenceAdjuster's axis edit) -------------

  private get limit(): number {
    return this.axis === 'row' ? SheetDimensions.maxRows : SheetDimensions.maxColumns;
  }

  /** Where a line lands, or null when deleted or pushed past the sheet edge (never clamped). */
  private mapCoord(coord: number): number | null {
    if (coord < this.at) return coord;
    if (this.operation === 'insert') {
      const shifted = coord + this.count;
      return shifted >= this.limit ? null : shifted;
    }
    return coord < this.at + this.count ? null : coord - this.count;
  }

  private mapSpan(min: number, max: number): [number, number] | null {
    const lo = this.mapCoord(min);
    if (lo === null) return null;
    return [lo, this.mapCoord(max) ?? this.limit - 1];
  }

  /** The first and last lines of [from, to] that survive a delete. */
  private survivingExtremes(from: number, to: number): [number, number] | null {
    if (this.operation === 'insert') return [from, to];
    const at = this.at;
    const deleteEnd = at + this.count;
    const min = from < at ? from : Math.max(from, deleteEnd) <= to ? Math.max(from, deleteEnd) : null;
    const max = to >= deleteEnd ? to : to < at ? to : at - 1 >= from ? at - 1 : null;
    return min !== null && max !== null ? [min, max] : null;
  }
}
