import type { CellRange } from '../../model/cellRange.ts';
import type { CellValue } from '../../values/cellValue.ts';
import type { Expr } from '../ast/expr.ts';
import type { CivilDateTime } from './dateConversion.ts';

/**
 * Everything the evaluator needs from the workbook: cell lookup (across sheets), named-range
 * resolution, the active locale, and a clock for the volatile date/time functions. Implemented over
 * a live workbook at the recalc layer; tests provide a lightweight fake.
 */
export interface EvaluationContext {
  /** The sheet a formula lives on; unqualified references resolve here. */
  readonly currentSheet: string;
  /**
   * The row and column of the formula's own cell. A range used where one value is needed takes
   * the cell in this row (or column), as in Calc; without it such a range is #VALUE!.
   */
  readonly formulaCell?: { readonly row: number; readonly column: number } | undefined;
  /** BCP-47 locale for VALUE/TEXT/DATE parsing and formatting. */
  readonly locale: string;
  /** Current local date/time for TODAY/NOW (injectable for tests). */
  now(): CivilDateTime;
  /** The displayed value of a cell. `sheet` null means `currentSheet`. */
  getCell(sheet: string | null, row: number, column: number): CellValue;
  /**
   * Whether a sheet exists (`null`: the current sheet). Leaving it out means every sheet does; a
   * reference to a sheet that does not is #REF!.
   */
  hasSheet?(sheet: string | null): boolean;
  /**
   * The names of the sheets from `first` to `last` in tab order, both included, for a range across
   * sheets (`Sheet1.A1:Sheet3.B2`). Null when either sheet does not exist; leaving it out makes
   * every such range #REF!.
   */
  sheetsBetween?(first: string, last: string): readonly string[] | null;
  /**
   * Why a row is hidden: `filtered` when a filter hid it, `hidden` when it was hidden by hand, null
   * when it shows. SUBTOTAL reads it; leaving it out means every row shows.
   */
  rowHiddenBy?(sheet: string | null, row: number): 'filtered' | 'hidden' | null;
  /** A cell's formula text, or null when it holds none. Leaving it out means no cell does. */
  cellFormula?(sheet: string | null, row: number, column: number): string | null;
  /** Resolve a defined name to its range; null for an unknown name or one that holds an expression. */
  resolveName(name: string): CellRange | null;
  /**
   * The parsed expression a name holds (`TaxRate` = `0.2`), or null when it holds a range or is
   * unknown. Leaving it out means no name holds an expression.
   */
  nameExpression?(name: string): Expr | null;
  /**
   * The last row and column of a sheet that can hold a value (null when it has none), so a range
   * like A:A reads only up to it. `undefined` (or leaving it out) means unknown: read every cell.
   */
  cellExtent?(
    sheet: string | null,
  ): { readonly lastRow: number; readonly lastColumn: number } | null | undefined;
}

/** The outcome of evaluating a formula plus whether it is volatile (TODAY/NOW). */
export interface EvaluationResult {
  readonly value: CellValue;
  readonly volatile: boolean;
}
