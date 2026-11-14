import { CellRange } from '../../model/cellRange.ts';
import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import type { EvaluationContext } from '../evaluation/evaluationContext.ts';

/** The rows of an inline array, all the same length. */
export type ArrayRows = readonly (readonly CellValue[])[];

/**
 * A read-only 2D view over a resolved range, fetched lazily from the evaluation context so a large
 * range is never materialized. Coordinates are zero-based within the range. It can also stand for an
 * inline array (`{1,2;3,4}`), whose `range` is then only its shape anchored at A1 (see `isArray`).
 */
export class RangeValues {
  readonly range: CellRange;
  readonly #context: EvaluationContext;
  readonly #sheet: string;
  readonly #array: ArrayRows | null;
  #reach: { rows: number; columns: number } | undefined;

  constructor(range: CellRange, context: EvaluationContext, array: ArrayRows | null = null) {
    this.range = range;
    this.#context = context;
    this.#sheet = range.sheet ?? context.currentSheet;
    this.#array = array;
  }

  /** A view over an inline array's values. */
  static ofArray(rows: ArrayRows, context: EvaluationContext): RangeValues {
    const columns = rows[0]?.length ?? 0;
    return new RangeValues(CellRange.fromBounds(0, 0, rows.length - 1, columns - 1, null), context, rows);
  }

  /** True for an inline array: its values are constants, and `range` is not cells on a sheet. */
  get isArray(): boolean {
    return this.#array !== null;
  }

  get rowCount(): number {
    return this.range.rowCount;
  }

  get columnCount(): number {
    return this.range.columnCount;
  }

  /**
   * How many rows and columns, counted from the range's top-left, can hold a value on its sheet;
   * past them every cell is empty. Not clamped to the range's own size (SUMIF reshapes its sum
   * range); Infinity when the context can't tell.
   */
  get reach(): { readonly rows: number; readonly columns: number } {
    if (!this.#reach) {
      const extent = this.#array ? undefined : this.#context.cellExtent?.(this.#sheet);
      this.#reach = this.#array
        ? { rows: this.rowCount, columns: this.columnCount }
        : extent === undefined
          ? { rows: Infinity, columns: Infinity }
          : extent === null
            ? { rows: 0, columns: 0 }
            : {
                rows: Math.max(0, extent.lastRow - this.range.top + 1),
                columns: Math.max(0, extent.lastColumn - this.range.left + 1),
              };
    }
    return this.#reach;
  }

  /** Rows of this range that can hold a value; the rest are empty. */
  get usedRowCount(): number {
    return Math.min(this.rowCount, this.reach.rows);
  }

  /** Columns of this range that can hold a value; the rest are empty. */
  get usedColumnCount(): number {
    return Math.min(this.columnCount, this.reach.columns);
  }

  get(rowWithin: number, columnWithin: number): CellValue {
    if (rowWithin < 0 || rowWithin >= this.rowCount || columnWithin < 0 || columnWithin >= this.columnCount) {
      return CellValue.error(CellError.reference('Index outside the range.'));
    }
    if (this.#array) return this.#array[rowWithin]![columnWithin]!;
    return this.#context.getCell(this.#sheet, this.range.top + rowWithin, this.range.left + columnWithin);
  }

  /** Every value in the range, row-major. */
  *values(): Generator<CellValue> {
    for (let r = 0; r < this.rowCount; r++) for (let c = 0; c < this.columnCount; c++) yield this.get(r, c);
  }

  /**
   * The values in row-major order, leaving out the empty cells past the sheet's last used row and
   * column. A whole-column range like A:A reads only as far as the data goes.
   */
  *usedValues(): Generator<CellValue> {
    const rows = this.usedRowCount;
    const columns = this.usedColumnCount;
    for (let r = 0; r < rows; r++) for (let c = 0; c < columns; c++) yield this.get(r, c);
  }
}
