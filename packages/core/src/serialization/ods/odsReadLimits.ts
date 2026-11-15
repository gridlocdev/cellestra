import { SheetDimensions } from '../../model/sheetDimensions.ts';
import type { XmlElement, XmlName } from '../../xml/index.ts';

export const MAX_COLUMNS = SheetDimensions.maxColumns;
export const MAX_ROWS = SheetDimensions.maxRows;

/** A styled blank run larger than this is applied as a row or column style instead of per cell. */
export const STYLED_RUN_CELL_LIMIT = 1 << 16;
/** The most cells one repeated run may create; past it the file is refused, not half-read. */
export const CONTENT_RUN_CELL_LIMIT = 1 << 22;
/**
 * The most cells repeated runs may create in one workbook, all sheets together: as many as one run
 * may cover, since a few rows of runs at that cap would fill memory from a file of a few hundred
 * bytes. LibreOffice repeats only identical neighbours (and Cellestra reads big blank styled runs
 * as row or column formats), so a real file stays far below it.
 */
export const REPEATED_CELL_BUDGET = CONTENT_RUN_CELL_LIMIT;
/**
 * The most merges repeated spanned cells may create in one workbook. LibreOffice never repeats a
 * cell that starts a merge, and neither does Cellestra, so only a corrupt or hostile file gets near.
 */
export const REPEATED_MERGE_BUDGET = 1 << 15;

/** The cells and merges repeats have created so far, shared by every sheet of a workbook. */
export interface RepeatBudget {
  cells: number;
  merges: number;
}

/** The width of a LibreOffice sheet before 7.4: a run that reaches it reaches the end of the row. */
export const LEGACY_COLUMNS = 1024;

/** Thrown when a file asks for more cells than can be held (a hostile or corrupt repeat count). */
export class OdsContentTooLargeError extends Error {
  override readonly name = 'OdsContentTooLargeError';
}

/** Count `count` more cells made by repeats on `sheet`; throws past the workbook's budget. */
export function chargeRepeatedCells(repeats: RepeatBudget, count: number, sheet: string): void {
  repeats.cells += count;
  if (repeats.cells > REPEATED_CELL_BUDGET) {
    throw new OdsContentTooLargeError(
      `Sheet '${sheet}' repeats cells past ${REPEATED_CELL_BUDGET} in the workbook, the most repeats may create.`,
    );
  }
}

/** `int.TryParse(…) && n > 0 ? n : 1`. */
export function readPositive(element: XmlElement, attribute: XmlName): number {
  const text = element.attr(attribute)?.trim() ?? '';
  if (!/^[+-]?\d+$/.test(text)) return 1;
  const n = Number(text);
  return n > 0 && n <= 2147483647 ? n : 1;
}
