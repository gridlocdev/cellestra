import type { Worksheet } from '../../model/worksheet.ts';
import { StylePool } from '../../styling/stylePool.ts';
import type { XmlElement } from '../../xml/index.ts';
import {
  chargeRepeatedCells,
  CONTENT_RUN_CELL_LIMIT,
  LEGACY_COLUMNS,
  MAX_COLUMNS,
  STYLED_RUN_CELL_LIMIT,
  type RepeatBudget,
} from './odsReadLimits.ts';
import type { OdsStyleIndex } from './odsStyleIndex.ts';
import type { StyleRegistry } from './odsStyleRegistry.ts';
import { Table } from './odsXmlNames.ts';

/**
 * Per-cell styles (`table:style-name` on a cell), and how big styled blank runs read: a run over a
 * whole row or down to the end of the sheet is a row or column format, not a million cells.
 */

// ── Reading ─────────────────────────────────────────────────────────────────

/** A cell element's own style, as a style id; null when it names none. */
export function cellStyleIdOf(element: XmlElement, styles: OdsStyleIndex): number | null {
  const name = element.attr(Table('style-name'));
  return name === null ? null : styles.cellStyleId(name);
}

/** A run of styled blank cells down to the end of the sheet, applied as a column format. */
export interface ColumnFormat {
  readonly from: number;
  readonly to: number;
  readonly styleId: number;
}

/** Where a run of cells read from one element goes. */
export interface StyledRun {
  readonly column: number;
  readonly repeat: number;
  readonly rowCount: number;
  /** The run holds a value, formula, link or kept markup (else it is styled blanks). */
  readonly hasContent: boolean;
  readonly styleId: number | null;
  /** It is the row element's last cell. */
  readonly lastInRow: boolean;
  /** The style every cell before it in the row has: undefined before the first, null once they differ. */
  readonly rowStyleBefore: number | null | undefined;
  /** The row element reaches the end of the sheet (or the table). */
  readonly trailingRows: boolean;
}

/** A styled blank run read as a format: the whole row's, or the columns'. */
export type RunFormat = 'row' | 'column' | null;

/**
 * Whether a run is a row or column format. A row or column format would style every cell outside
 * the run too, so a run only becomes one when it covers the whole row or column. A run too big to
 * hold as cells is a format anyway, as it always was: a lossy read beats refusing the file.
 */
export function runFormat(run: StyledRun): RunFormat {
  if (run.hasContent) return null;
  const area = run.repeat * run.rowCount;
  const big = area > STYLED_RUN_CELL_LIMIT;
  const tooBigForCells = area > CONTENT_RUN_CELL_LIMIT;
  const reachesRowEnd = run.column + run.repeat >= LEGACY_COLUMNS;
  const rowFormat =
    run.lastInRow &&
    (big || (run.column + run.repeat >= MAX_COLUMNS && run.repeat > 1)) &&
    ((reachesRowEnd && (run.rowStyleBefore === undefined || run.rowStyleBefore === run.styleId)) ||
      tooBigForCells);
  if (rowFormat) return 'row';
  return big && run.trailingRows ? 'column' : null;
}

/** A cell above a column format and the style it keeps. */
export interface Pin {
  readonly row: number;
  readonly column: number;
  readonly styleId: number;
}

/**
 * The cells of `column … column + width - 1` above `row` that a column format of `styleId` would
 * restyle, each with the style it has now; null when there are more than `STYLED_RUN_CELL_LIMIT`
 * (then the run is cheaper to hold as cells) or more than `CONTENT_RUN_CELL_LIMIT` to look at. A cell or row with a style of its own keeps it
 * anyway, so in the usual layout (a styled header row over formatted columns) there is nothing to
 * pin; a header in the default style is pinned to a plain style of its own.
 */
export function pinsAbove(
  sheet: Worksheet,
  row: number,
  column: number,
  width: number,
  styleId: number,
  styles: OdsStyleIndex,
): Pin[] | null {
  if (row * width > CONTENT_RUN_CELL_LIMIT) return null;
  const pins: Pin[] = [];
  for (let r = 0; r < row; r++) {
    if ((sheet.findRow(r)?.styleId ?? StylePool.defaultId) !== StylePool.defaultId) continue;
    for (let c = column; c < column + width; c++) {
      if ((sheet.find(r, c)?.styleId ?? StylePool.defaultId) !== StylePool.defaultId) continue;
      const now = sheet.findColumn(c)?.styleId ?? StylePool.defaultId;
      if (now === styleId) continue;
      if (pins.length === STYLED_RUN_CELL_LIMIT) return null;
      pins.push({
        row: r,
        column: c,
        styleId: now === StylePool.defaultId ? styles.plainCellStyleId() : now,
      });
    }
  }
  return pins;
}

/** Give each pinned cell its style; the new cells count against the workbook's repeat budget. */
export function pinCells(sheet: Worksheet, pins: readonly Pin[], repeats: RepeatBudget): void {
  chargeRepeatedCells(repeats, pins.length, sheet.name);
  for (const pin of pins) sheet.getOrCreate(pin.row, pin.column).styleId = pin.styleId;
}

// ── Writing ─────────────────────────────────────────────────────────────────

/** Name a cell's own style on its element (the default style is left unnamed). */
export function writeCellStyle(element: XmlElement, styleId: number, registry: StyleRegistry): void {
  if (styleId !== StylePool.defaultId) element.setAttr(Table('style-name'), registry.cellStyle(styleId));
}
