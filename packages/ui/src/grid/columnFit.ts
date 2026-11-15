import {
  NumberFormatter,
  type CellStyle,
  type CellValue,
  type Workbook,
  type Worksheet,
} from '@cellestra/core';
import { fittedWidth, type TextMeasure } from './cellText.ts';

/** The widest a column gets when a CSV is opened or imported, so one long cell can't take the screen. */
export const CSV_FIT_MAX_WIDTH = 400;
/** A CSV's columns are fitted to this many first rows, so a large file still opens quickly. */
export const CSV_FIT_ROWS = 2_000;

export interface ColumnFitOptions {
  /** Which columns to fit; all by default. */
  readonly columns?: (column: number) => boolean;
  /** Only rows above this one are measured; all by default. */
  readonly rowLimit?: number;
  /** The widest a fitted column gets; no limit by default. */
  readonly maxWidth?: number;
}

/**
 * The width each column needs to show its widest value whole, measured the way the grid draws it
 * (`fittedWidth`), rounded up to whole px. Columns with nothing to show are left out. One pass over
 * the stored cells, with each style's texts measured once: a column of repeated values is cheap.
 */
export function fittedColumnWidths(
  ctx: TextMeasure,
  workbook: Workbook,
  sheet: Worksheet,
  display: (value: CellValue, style: CellStyle) => string,
  options: ColumnFitOptions = {},
): Map<number, number> {
  const { columns, rowLimit = Infinity, maxWidth = Infinity } = options;
  const measured = new Map<string, number>();
  const widths = new Map<number, number>();
  for (const [row, column, cell] of sheet.enumerateCells()) {
    if (row >= rowLimit || (columns && !columns(column))) continue;
    if (cell.value.kind === 'empty' || sheet.findMerge(row, column)) continue;
    const styleId = sheet.effectiveStyleId(row, column);
    const style = workbook.styles.get(styleId);
    const text = display(cell.value, style);
    const key = `${styleId}\u0000${text}`;
    let width = measured.get(key);
    if (width === undefined) {
      // Wrapped flat text folds at the column's edge, so it asks for no width.
      width = fittedWidth(ctx, text, style);
      measured.set(key, width);
    }
    if (width > (widths.get(column) ?? 0)) widths.set(column, width);
  }
  for (const [column, width] of widths) widths.set(column, Math.min(maxWidth, Math.ceil(width)));
  return widths;
}

/**
 * Fit a CSV sheet's used columns to their contents, in place: no edit, no undo step. Widths are
 * capped at `CSV_FIT_MAX_WIDTH`, and only the first `CSV_FIT_ROWS` rows are measured.
 */
export function fitCsvColumns(ctx: TextMeasure, workbook: Workbook, sheet: Worksheet): void {
  const locale = workbook.settings.effectiveLocale;
  const display = (value: CellValue, style: CellStyle) =>
    NumberFormatter.format(value, style.numberFormat.code, locale);
  const widths = fittedColumnWidths(ctx, workbook, sheet, display, {
    rowLimit: CSV_FIT_ROWS,
    maxWidth: CSV_FIT_MAX_WIDTH,
  });
  for (const [column, width] of widths) sheet.getOrCreateColumn(column).width = width;
}

/** A canvas to measure text with, or null where there is no DOM (tests, workers). */
export function canvasMeasure(): TextMeasure | null {
  if (typeof document === 'undefined' || typeof document.createElement !== 'function') return null;
  return document.createElement('canvas').getContext('2d');
}
