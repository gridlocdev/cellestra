import type { CellRange } from '../../model/cellRange.ts';
import type { CellValue } from '../../values/cellValue.ts';
import type { CivilDateTime } from './dateConversion.ts';

/**
 * Everything the evaluator needs from the workbook: cell lookup (across sheets), named-range
 * resolution, the active locale, and a clock for the volatile date/time functions. Implemented over
 * a live workbook at the recalc layer; tests provide a lightweight fake.
 */
export interface EvaluationContext {
  /** The sheet a formula lives on; unqualified references resolve here. */
  readonly currentSheet: string;
  /** BCP-47 locale for VALUE/TEXT/DATE parsing and formatting. */
  readonly locale: string;
  /** Current local date/time for TODAY/NOW (injectable for tests). */
  now(): CivilDateTime;
  /** The displayed value of a cell. `sheet` null means `currentSheet`. */
  getCell(sheet: string | null, row: number, column: number): CellValue;
  /** Resolve a defined name to its range; null for an unknown name. */
  resolveName(name: string): CellRange | null;
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
