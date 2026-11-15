import { Validator, type RangeResolver } from '../data/dataValidation.ts';
import type { DataValidation } from '../data/types.ts';
import type { Workbook } from '../model/workbook.ts';
import { CompositeEditCommand, requireSheet, type EditCommand } from './editCommand.ts';
import { EditCommands } from './editCommands.ts';
import { InputParser } from './inputParser.ts';
import { TableCommands } from './tableCommands.ts';

/** Typing text into a cell: the edit, and the validation rule it breaks, if any. */
export interface CellEntryPlan {
  readonly command: EditCommand;
  /**
   * The cell's rule when the typed value breaks it, else null. With a `stop` response the entry
   * must not go ahead; the others only warn. A formula is never checked (its value isn't known yet).
   */
  readonly brokenRule: DataValidation | null;
}

/**
 * What entering `text` at (row, column) does: the same parse as typing (formulas, numbers, dates
 * with their format), and a table just beside the cell grows to take it in, in the same undo step.
 * `resolve` reads a list rule's source range; without it a range-sourced list allows nothing.
 */
export function planCellEntry(
  workbook: Workbook,
  sheet: string,
  row: number,
  column: number,
  text: string,
  locale: string,
  resolve: RangeResolver | null = null,
): CellEntryPlan {
  const ws = requireSheet(workbook, sheet);
  const rule = ws.findValidation(row, column) ?? null;
  let brokenRule: DataValidation | null = null;
  if (rule) {
    const parsed = InputParser.parse(text, locale);
    if (parsed.formula === null && !Validator.isValid(rule, parsed.value, locale, resolve)) brokenRule = rule;
  }
  const set = EditCommands.setInput(sheet, row, column, text, locale, workbook);
  if (text === '' || ws.findTable(row, column)) return { command: set, brokenRule };
  for (const t of ws.tables) {
    const expand = TableCommands.expandTableForEdit(workbook, sheet, t, row, column);
    if (expand) return { command: new CompositeEditCommand(set.label, [set, expand]), brokenRule };
  }
  return { command: set, brokenRule };
}
