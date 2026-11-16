import {
  CellAddress,
  CellRange,
  ClipboardBlock,
  DataEditCommands,
  FillCommand,
  FormatBlock,
  PasteBlockCommand,
  SheetDimensions,
  ValidationPreflight,
  type CellValue,
  type ColumnFilter,
  type IncomingWrite,
  type PasteMode,
  type Worksheet,
} from '@cellestra/core';
import type { EditSession } from '../infra/editSession.ts';
import { CliError } from '../infra/errors.ts';
import { address } from '../infra/report.ts';

/**
 * The block ops behind `copy`, `move`, `fill`, `paint`, and `filter` — the desktop's clipboard,
 * fill handle, format painter, and AutoFilter, driven by the same core commands (so a copied
 * formula's relative references shift exactly as in the app, and a Stop validation rule aborts a
 * paste the same way).
 */

export function copyOrMove(
  session: EditSession,
  fromSheet: Worksheet,
  from: CellRange,
  toSheet: Worksheet,
  to: CellRange,
  cut: boolean,
  modeText: string,
): void {
  let mode: PasteMode;
  switch (modeText.trim().toLowerCase()) {
    case 'all':
      mode = 'all';
      break;
    case 'values':
      mode = 'values';
      break;
    case 'formats':
    case 'formatting':
      mode = 'formats';
      break;
    case 'formulas':
      mode = 'formulas';
      break;
    default:
      throw new CliError('"mode" must be all, values, formats, or formulas.');
  }
  if (cut && mode !== 'all') throw new CliError('a move always carries everything; "mode" applies to copy.');
  const block = ClipboardBlock.capture(fromSheet, from, cut);
  const paste = new PasteBlockCommand(toSheet.name, to, block, mode);
  preflight(session, toSheet, paste.projectWrites());
  session.apply(paste);
}

export function fill(
  session: EditSession,
  sheet: Worksheet,
  range: CellRange,
  source: CellRange | null,
): void {
  const seed =
    source ??
    (range.rowCount >= range.columnCount
      ? new CellRange(new CellAddress(range.left, range.top), new CellAddress(range.right, range.top))
      : new CellRange(new CellAddress(range.left, range.top), new CellAddress(range.left, range.bottom)));
  if (!range.contains(seed.top, seed.left) || !range.contains(seed.bottom, seed.right)) {
    throw new CliError(
      `the source ${seed.toString()} must lie inside the range ${range.toString()} being filled.`,
    );
  }
  const command = new FillCommand(sheet.name, seed, range);
  preflight(session, sheet, command.projectWrites(session.workbook));
  session.apply(command);
}

export function paint(
  session: EditSession,
  fromSheet: Worksheet,
  from: CellRange,
  toSheet: Worksheet,
  to: CellRange,
): void {
  if (FormatBlock.tooLarge(fromSheet, from))
    throw new CliError(`${from.toString()} is too big to copy formats from; give a smaller range.`);
  const block = FormatBlock.capture(session.workbook, fromSheet, from);
  session.apply(block.applyTo(toSheet.name, to));
  for (const [row, column] of to.positions()) session.touched(toSheet.name, row, column);
}

export function filter(
  session: EditSession,
  sheet: Worksheet,
  range: CellRange,
  header: boolean,
  clear: boolean,
  columns: ReadonlyMap<number, ColumnFilter>,
): void {
  if (clear) {
    session.apply(DataEditCommands.clearFilter(sheet.name, range));
    return;
  }
  if (columns.size === 0) {
    throw new CliError(
      'a filter needs at least one column criterion ("columns": {"B": {"condition": ">5"}}), or "clear": true.',
    );
  }
  for (const column of columns.keys()) {
    if (column < range.left || column > range.right) {
      throw new CliError(`column ${SheetDimensions.columnToLetters(column)} is outside ${range.toString()}.`);
    }
  }
  session.apply(
    DataEditCommands.applyFilter(sheet.name, range, { columns, hasHeader: header }, session.locale),
  );
  let hidden = 0;
  for (let row = range.top + (header ? 1 : 0); row <= range.bottom; row++) {
    if (sheet.isRowHidden(row)) hidden++;
  }
  session.field('hiddenRows', hidden);
  session.note(`Filtered ${range.toString()}: ${hidden} row(s) hidden.`);
}

/** The app's rule: a paste or fill that would violate a Stop validation rule doesn't happen at all. */
function preflight(session: EditSession, sheet: Worksheet, writes: Iterable<IncomingWrite>): void {
  const violations = ValidationPreflight.findStopViolations(sheet, writes, session.locale, (range) =>
    resolve(sheet, range),
  );
  const first = violations[0];
  if (first) {
    const [row, column] = first;
    throw new CliError(
      `${address(sheet.name, row, column)} would violate its validation rule (${violations.length} cell(s) in all); nothing was written.`,
    );
  }
}

function* resolve(sheet: Worksheet, range: CellRange): Iterable<CellValue> {
  for (const [row, column] of range.positions()) yield sheet.getValue(row, column);
}
