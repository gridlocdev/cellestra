import { DataToolCommands } from './dataToolCommands.ts';
import { SheetLayoutCommands } from './sheetLayoutCommands.ts';
import { SortFilterCommands } from './sortFilterCommands.ts';
import { TableCommands } from './tableCommands.ts';
import { ValidationCommands } from './validationCommands.ts';

/**
 * Undoable data-menu edits in one place: sort and filter, duplicates and text to columns, tables,
 * validation, and layout resets. Each feature lives in its own module.
 */
export const DataEditCommands = {
  ...SortFilterCommands,
  ...DataToolCommands,
  ...TableCommands,
  ...ValidationCommands,
  ...SheetLayoutCommands,
} as const;
