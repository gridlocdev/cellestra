import type { CellRange } from '../../model/cellRange.ts';
import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import type { EvaluationContext } from '../evaluation/evaluationContext.ts';

/**
 * A read-only 2D view over a resolved range, fetched lazily from the evaluation context so a large
 * range is never materialized. Coordinates are zero-based within the range.
 */
export class RangeValues {
  readonly range: CellRange;
  readonly #context: EvaluationContext;
  readonly #sheet: string;

  constructor(range: CellRange, context: EvaluationContext) {
    this.range = range;
    this.#context = context;
    this.#sheet = range.sheet ?? context.currentSheet;
  }

  get rowCount(): number {
    return this.range.rowCount;
  }

  get columnCount(): number {
    return this.range.columnCount;
  }

  get(rowWithin: number, columnWithin: number): CellValue {
    if (rowWithin < 0 || rowWithin >= this.rowCount || columnWithin < 0 || columnWithin >= this.columnCount) {
      return CellValue.error(CellError.reference('Index outside the range.'));
    }
    return this.#context.getCell(this.#sheet, this.range.top + rowWithin, this.range.left + columnWithin);
  }

  /** Every value in the range, row-major. */
  *values(): Generator<CellValue> {
    for (let r = 0; r < this.rowCount; r++) for (let c = 0; c < this.columnCount; c++) yield this.get(r, c);
  }
}
