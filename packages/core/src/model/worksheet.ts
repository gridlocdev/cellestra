import type { ChartDefinition } from '@cellestra/charts';
import { cloneChart } from '@cellestra/charts';
import { CellValue } from '../values/cellValue.ts';
import type { DataValidation, FilterSpec, SortSpec, Table } from '../data/types.ts';
import { Cell } from './cell.ts';
import { CellAddress } from './cellAddress.ts';
import { CellRange } from './cellRange.ts';
import { bumpLineLayout, Column, lineLayoutGeneration, Row } from './lines.ts';
import type { MergedRegion } from './mergedRegion.ts';
import { cellKey, keyColumn, keyRow, SheetDimensions } from './sheetDimensions.ts';
import { Defaults } from '../defaults.ts';
import type { PageSetup } from './pageSetup.ts';

const DEFAULT_STYLE_ID = 0;

/** A snapshot of a sheet's line layout, merges, tables and panes (not its cells), for structural undo. */
export interface WorksheetLayout {
  readonly rows: ReadonlyMap<number, Row>;
  readonly columns: ReadonlyMap<number, Column>;
  readonly merges: readonly MergedRegion[];
  readonly tables: readonly Table[];
  readonly frozenRows: number;
  readonly frozenColumns: number;
}

/**
 * One sheet: sparse cells keyed by packed (row, column), sparse row/column layout, merges,
 * frozen panes, tables, filter/sort state, and validation rules. A chart sheet carries a `chart`.
 */
export class Worksheet {
  private readonly cells = new Map<number, Cell>();
  private readonly rowMap = new Map<number, Row>();
  private readonly columnMap = new Map<number, Column>();
  private mergedRegions: MergedRegion[] = [];
  private tableList: Table[] = [];
  private validationList: DataValidation[] = [];
  /** Last row/column holding a stored cell; `undefined` when it must be recomputed. */
  #cellExtent: { lastRow: number; lastColumn: number } | null | undefined = null;

  name: string;
  chart: ChartDefinition | null = null;
  frozenRows = 0;
  frozenColumns = 0;
  activeFilter: CellRange | null = null;
  activeFilterSpec: FilterSpec | null = null;
  activeSort: SortSpec | null = null;
  /** How the sheet prints; null for the default page (A4, as LibreOffice). */
  pageSetup: PageSetup | null = null;

  constructor(name: string) {
    this.name = name;
  }

  get isChartSheet(): boolean {
    return this.chart !== null;
  }

  get cellCount(): number {
    return this.cells.size;
  }

  // ---- Cells ------------------------------------------------------------

  find(row: number, column: number): Cell | undefined {
    return this.cells.get(cellKey(row, column));
  }

  findAt(address: CellAddress): Cell | undefined {
    return this.find(address.row, address.column);
  }

  getValue(row: number, column: number): CellValue {
    return this.find(row, column)?.value ?? CellValue.empty;
  }

  getOrCreate(row: number, column: number): Cell {
    validatePosition(row, column);
    const key = cellKey(row, column);
    let cell = this.cells.get(key);
    if (!cell) {
      cell = new Cell();
      this.cells.set(key, cell);
      this.#grew(row, column);
    }
    return cell;
  }

  /** Set a literal. Setting empty clears content but keeps style/hyperlink if present. */
  setValue(row: number, column: number, value: CellValue): void {
    if (value.kind === 'empty') {
      const existing = this.find(row, column);
      if (!existing) return;
      existing.formula = null;
      existing.value = CellValue.empty;
      this.pruneIfBlank(row, column, existing);
      return;
    }
    const cell = this.getOrCreate(row, column);
    cell.formula = null;
    cell.value = value;
  }

  /** Formula source without the leading `=`. */
  setFormula(row: number, column: number, formula: string, cachedValue: CellValue = CellValue.empty): void {
    const cell = this.getOrCreate(row, column);
    cell.formula = formula;
    cell.value = cachedValue;
  }

  remove(row: number, column: number): void {
    if (this.cells.delete(cellKey(row, column))) this.#shrank(row, column);
  }

  /** Put a cell object at a position (or remove when null/blank). */
  putCell(row: number, column: number, cell: Cell | null | undefined): void {
    if (!cell || cell.isBlank) this.remove(row, column);
    else {
      this.cells.set(cellKey(row, column), cell);
      this.#grew(row, column);
    }
  }

  pruneIfBlank(row: number, column: number, cell: Cell): void {
    if (cell.isBlank) this.remove(row, column);
  }

  /**
   * The last row and column that hold a stored cell (null for no cells). Cached, so formulas can
   * clip whole-column and whole-row ranges cheaply; cells past it all read as empty.
   */
  cellExtent(): { readonly lastRow: number; readonly lastColumn: number } | null {
    if (this.#cellExtent === undefined) {
      let lastRow = -1;
      let lastColumn = -1;
      for (const key of this.cells.keys()) {
        lastRow = Math.max(lastRow, keyRow(key));
        lastColumn = Math.max(lastColumn, keyColumn(key));
      }
      this.#cellExtent = lastRow < 0 ? null : { lastRow, lastColumn };
    }
    return this.#cellExtent;
  }

  #grew(row: number, column: number): void {
    const e = this.#cellExtent;
    if (e === undefined) return;
    if (e === null) this.#cellExtent = { lastRow: row, lastColumn: column };
    else if (row > e.lastRow || column > e.lastColumn) {
      this.#cellExtent = { lastRow: Math.max(row, e.lastRow), lastColumn: Math.max(column, e.lastColumn) };
    }
  }

  #shrank(row: number, column: number): void {
    const e = this.#cellExtent;
    if (e && (row === e.lastRow || column === e.lastColumn)) this.#cellExtent = undefined;
  }

  /**
   * The stored cells inside `range`, found by looking up each position or by scanning the storage,
   * whichever is less work, so a huge, mostly empty range costs only its stored cells.
   */
  cellsIn(range: CellRange): [row: number, column: number, cell: Cell][] {
    const out: [number, number, Cell][] = [];
    if (range.cellCount <= this.cells.size) {
      for (const [row, column] of range.positions()) {
        const cell = this.find(row, column);
        if (cell) out.push([row, column, cell]);
      }
      return out;
    }
    for (const [key, cell] of this.cells) {
      const row = keyRow(key);
      const column = keyColumn(key);
      if (range.contains(row, column)) out.push([row, column, cell]);
    }
    return out;
  }

  /** Every stored cell, in insertion order. Don't mutate the map while iterating. */
  *enumerateCells(): IterableIterator<[row: number, column: number, cell: Cell]> {
    for (const [key, cell] of this.cells) yield [keyRow(key), keyColumn(key), cell];
  }

  // ---- Rows & columns ---------------------------------------------------

  findRow(row: number): Row | undefined {
    return this.rowMap.get(row);
  }
  findColumn(column: number): Column | undefined {
    return this.columnMap.get(column);
  }

  getOrCreateRow(row: number): Row {
    let r = this.rowMap.get(row);
    if (!r) {
      r = new Row();
      this.rowMap.set(row, r);
    }
    return r;
  }

  getOrCreateColumn(column: number): Column {
    let c = this.columnMap.get(column);
    if (!c) {
      c = new Column();
      this.columnMap.set(column, c);
    }
    return c;
  }

  getRowHeight(row: number): number {
    return this.rowMap.get(row)?.height ?? Defaults.rowHeight;
  }
  getColumnWidth(column: number): number {
    return this.columnMap.get(column)?.width ?? Defaults.columnWidth;
  }
  /** Not shown, whether hidden by hand or by a filter. */
  isRowHidden(row: number): boolean {
    return this.rowMap.get(row)?.hidden ?? false;
  }
  /** Hidden by a filter (it may be hidden by hand as well). */
  isRowFiltered(row: number): boolean {
    return this.rowMap.get(row)?.filtered ?? false;
  }
  isColumnHidden(column: number): boolean {
    return this.columnMap.get(column)?.hidden ?? false;
  }

  /**
   * Moves whenever row or column layout may have changed (it is shared by every sheet, so it can
   * also move for another sheet's change). Compare it to skip re-reading `rows` and `columns`.
   */
  get layoutStamp(): number {
    return lineLayoutGeneration();
  }

  get rows(): ReadonlyMap<number, Row> {
    return this.rowMap;
  }
  get columns(): ReadonlyMap<number, Column> {
    return this.columnMap;
  }

  /** Cell style, else row style, else column style, else default. */
  effectiveStyleId(row: number, column: number): number {
    const own = this.find(row, column)?.styleId ?? DEFAULT_STYLE_ID;
    if (own !== DEFAULT_STYLE_ID) return own;
    const rowStyle = this.rowMap.get(row)?.styleId ?? DEFAULT_STYLE_ID;
    if (rowStyle !== DEFAULT_STYLE_ID) return rowStyle;
    return this.columnMap.get(column)?.styleId ?? DEFAULT_STYLE_ID;
  }

  // ---- Merged regions ---------------------------------------------------

  get merges(): readonly MergedRegion[] {
    return this.mergedRegions;
  }

  findMerge(row: number, column: number): MergedRegion | undefined {
    return this.mergedRegions.find((m) => m.covers(row, column));
  }

  addMerge(region: MergedRegion): void {
    if (this.mergedRegions.some((m) => m.range.intersects(region.range)))
      throw new Error('Merged regions may not overlap (PRD §8.10).');
    this.mergedRegions.push(region);
  }

  removeMerge(row: number, column: number): boolean {
    const region = this.findMerge(row, column);
    if (!region) return false;
    this.mergedRegions = this.mergedRegions.filter((m) => m !== region);
    return true;
  }

  replaceMerges(regions: Iterable<MergedRegion>): void {
    this.mergedRegions = [...regions];
  }

  /** The range a click at (row, column) selects: the whole merge, or the single cell. */
  selectionRangeAt(row: number, column: number): CellRange {
    return this.findMerge(row, column)?.range ?? new CellRange(new CellAddress(column, row));
  }

  /** One keyboard step that hops over merges and snaps onto a merge's anchor. */
  stepOverMerges(row: number, column: number, rowDelta: number, columnDelta: number): [number, number] {
    const from = this.findMerge(row, column);
    if (from) {
      if (rowDelta > 0) row = from.range.bottom;
      else if (rowDelta < 0) row = from.range.top;
      if (columnDelta > 0) column = from.range.right;
      else if (columnDelta < 0) column = from.range.left;
    }
    let r = clamp(row + rowDelta, 0, SheetDimensions.maxRows - 1);
    let c = clamp(column + columnDelta, 0, SheetDimensions.maxColumns - 1);
    const into = this.findMerge(r, c);
    if (into) {
      r = into.range.top;
      c = into.range.left;
    }
    return [r, c];
  }

  // ---- Tables & validation ----------------------------------------------

  get tables(): readonly Table[] {
    return this.tableList;
  }
  get validations(): readonly DataValidation[] {
    return this.validationList;
  }

  addTable(t: Table): void {
    if (this.tableList.some((x) => x.range.intersects(t.range)))
      throw new Error('Tables may not overlap (PRD §8.10).');
    this.tableList.push(t);
  }

  removeTable(t: Table): boolean {
    const before = this.tableList.length;
    this.tableList = this.tableList.filter((x) => x !== t);
    return this.tableList.length !== before;
  }

  replaceTable(existing: Table, replacement: Table): void {
    const i = this.tableList.indexOf(existing);
    if (i >= 0) this.tableList[i] = replacement;
  }

  findTable(row: number, column: number): Table | undefined {
    return this.tableList.find((t) => t.range.contains(row, column));
  }

  addValidation(v: DataValidation): void {
    this.validationList.push(v);
  }

  /** Swap in a whole new list of rules (reference fix-ups and their undo). */
  replaceValidations(validations: Iterable<DataValidation>): void {
    this.validationList = [...validations];
  }

  removeValidation(v: DataValidation): boolean {
    const before = this.validationList.length;
    this.validationList = this.validationList.filter((x) => x !== v);
    return this.validationList.length !== before;
  }

  findValidation(row: number, column: number): DataValidation | undefined {
    return this.validationList.find((v) => v.target.contains(row, column));
  }

  // ---- Extents ----------------------------------------------------------

  /**
   * The bounding box of everything that carries content or non-default layout. Rows and columns
   * are tracked independently; default-valued row/column entries (left by mere lookups) don't count.
   */
  getUsedRange(): CellRange | null {
    let hasRow = false;
    let hasCol = false;
    let top = Infinity;
    let bottom = 0;
    let left = Infinity;
    let right = 0;
    const rows = (from: number, to: number) => {
      hasRow = true;
      top = Math.min(top, from);
      bottom = Math.max(bottom, to);
    };
    const cols = (from: number, to: number) => {
      hasCol = true;
      left = Math.min(left, from);
      right = Math.max(right, to);
    };
    for (const key of this.cells.keys()) {
      const r = keyRow(key);
      const c = keyColumn(key);
      rows(r, r);
      cols(c, c);
    }
    for (const m of this.mergedRegions) {
      rows(m.range.top, m.range.bottom);
      cols(m.range.left, m.range.right);
    }
    for (const [r, model] of this.rowMap) if (!model.isDefault) rows(r, r);
    for (const [c, model] of this.columnMap) if (!model.isDefault) cols(c, c);
    if (!hasRow && !hasCol) return null;
    if (!hasRow) top = bottom = 0;
    if (!hasCol) left = right = 0;
    return CellRange.fromBounds(top, left, bottom, right);
  }

  clone(newName: string): Worksheet {
    const copy = new Worksheet(newName);
    for (const [k, c] of this.cells) copy.cells.set(k, c.clone());
    copy.#cellExtent = this.#cellExtent;
    for (const [k, r] of this.rowMap) copy.rowMap.set(k, r.clone());
    for (const [k, c] of this.columnMap) copy.columnMap.set(k, c.clone());
    // Merges, tables, validations are immutable — sharing them is a full copy.
    copy.mergedRegions = [...this.mergedRegions];
    copy.tableList = [...this.tableList];
    copy.validationList = [...this.validationList];
    copy.frozenRows = this.frozenRows;
    copy.frozenColumns = this.frozenColumns;
    copy.activeFilter = this.activeFilter;
    copy.activeFilterSpec = this.activeFilterSpec;
    copy.activeSort = this.activeSort;
    copy.pageSetup = this.pageSetup;
    copy.chart = this.chart ? cloneChart(this.chart) : null;
    return copy;
  }

  // ---- Structural edits ---------------------------------------------------

  captureLayout(): WorksheetLayout {
    return {
      rows: new Map([...this.rowMap].map(([k, r]) => [k, r.clone()])),
      columns: new Map([...this.columnMap].map(([k, c]) => [k, c.clone()])),
      merges: [...this.mergedRegions],
      tables: [...this.tableList],
      frozenRows: this.frozenRows,
      frozenColumns: this.frozenColumns,
    };
  }

  restoreLayout(layout: WorksheetLayout): void {
    bumpLineLayout();
    this.rowMap.clear();
    for (const [k, r] of layout.rows) this.rowMap.set(k, r.clone());
    this.columnMap.clear();
    for (const [k, c] of layout.columns) this.columnMap.set(k, c.clone());
    this.mergedRegions = [...layout.merges];
    this.tableList = [...layout.tables];
    this.frozenRows = layout.frozenRows;
    this.frozenColumns = layout.frozenColumns;
  }

  /**
   * Move every cell to the key `move` gives its position, keeping the same cell objects (nothing is
   * copied). A cell `move` maps to null is taken out; the removed cells come back by their old key.
   * `move` must not send two cells to one key.
   */
  moveCells(move: (row: number, column: number) => number | null): Map<number, Cell> {
    const moved = new Map<number, Cell>();
    const removed = new Map<number, Cell>();
    for (const [key, cell] of this.cells) {
      const target = move(keyRow(key), keyColumn(key));
      if (target === null) removed.set(key, cell);
      else moved.set(target, cell);
    }
    this.cells.clear();
    for (const [key, cell] of moved) this.cells.set(key, cell);
    this.#cellExtent = undefined;
    return removed;
  }
}

function clamp(v: number, lo: number, hi: number): number {
  return Math.min(hi, Math.max(lo, v));
}

function validatePosition(row: number, column: number): void {
  if (!SheetDimensions.isValidRow(row)) throw new RangeError(`Row index out of range: ${row}`);
  if (!SheetDimensions.isValidColumn(column)) throw new RangeError(`Column index out of range: ${column}`);
}
