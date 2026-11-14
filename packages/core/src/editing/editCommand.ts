import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';

/** A cell an edit touched, by sheet name: what recalculation and the grid repaint start from. */
export interface ChangedCell {
  readonly sheet: string;
  readonly row: number;
  readonly column: number;
}

/** A zero-based (row, column) position. */
export type Position = readonly [row: number, column: number];

/** A run of whole rows or whole columns of one sheet, `first` to `last` inclusive. */
export interface ChangedLines {
  readonly sheet: string;
  readonly first: number;
  readonly last: number;
}

/**
 * What an apply/undo changed: the cells whose content moved (for targeted recalc), or `structural`
 * when rows/columns/sheets shifted and everything must be re-planned. `changedRows` and
 * `changedColumns` are the whole lines whose cells may now look or measure differently without
 * their content moving (a line style, a width or height, a merge), so a view re-measures the cells
 * in those lines and leaves the rest.
 */
export interface EditResult {
  readonly changedCells: readonly ChangedCell[];
  readonly changedRows: readonly ChangedLines[];
  readonly changedColumns: readonly ChangedLines[];
  readonly structural: boolean;
}

const NO_CELLS: readonly ChangedCell[] = Object.freeze([]);
const NO_LINES: readonly ChangedLines[] = Object.freeze([]);
const result = (
  changedCells: readonly ChangedCell[],
  changedRows: readonly ChangedLines[],
  changedColumns: readonly ChangedLines[],
  structural: boolean,
): EditResult => ({ changedCells, changedRows, changedColumns, structural });

export const EditResult = {
  none: Object.freeze(result(NO_CELLS, NO_LINES, NO_LINES, false)),
  structuralChange: Object.freeze(result(NO_CELLS, NO_LINES, NO_LINES, true)),
  cells: (changedCells: readonly ChangedCell[]): EditResult =>
    result(changedCells, NO_LINES, NO_LINES, false),
  /** Rows `first` to `last` of `sheet` changed how their cells look or measure. */
  rows: (sheet: string, first: number, last: number): EditResult =>
    result(NO_CELLS, [{ sheet, first, last }], NO_LINES, false),
  /** Columns `first` to `last` of `sheet` changed how their cells look or measure. */
  columns: (sheet: string, first: number, last: number): EditResult =>
    result(NO_CELLS, NO_LINES, [{ sheet, first, last }], false),
} as const;

/**
 * One undoable edit. Commands address sheets by name so they survive a sheet object being swapped,
 * and each `apply` re-captures whatever its `undo` needs (redo re-applies from scratch).
 */
export interface EditCommand {
  readonly label: string;
  apply(workbook: Workbook): EditResult;
  undo(workbook: Workbook): EditResult;
}

/**
 * Thrown by a command's `apply` when it refuses an edit it can't carry out faithfully (sorting
 * across merged cells, inserting lines that would push data off the sheet). The workbook is left as
 * it was and nothing goes on the undo stack; `message` is written for the user.
 */
export class EditRefusedError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'EditRefusedError';
  }
}

/** A command from a pair of closures; state shared between them lives in the closure scope. */
export class ActionCommand implements EditCommand {
  readonly label: string;
  private readonly applyFn: (workbook: Workbook) => EditResult;
  private readonly undoFn: (workbook: Workbook) => EditResult;

  constructor(
    label: string,
    apply: (workbook: Workbook) => EditResult,
    undo: (workbook: Workbook) => EditResult,
  ) {
    this.label = label;
    this.applyFn = apply;
    this.undoFn = undo;
  }

  apply(workbook: Workbook): EditResult {
    return this.applyFn(workbook);
  }

  undo(workbook: Workbook): EditResult {
    return this.undoFn(workbook);
  }
}

/** Several commands as one undo step: applied in order, undone in reverse. */
export class CompositeEditCommand implements EditCommand {
  readonly label: string;
  readonly commands: readonly EditCommand[];

  constructor(label: string, commands: readonly EditCommand[]) {
    this.label = label;
    this.commands = commands;
  }

  apply(workbook: Workbook): EditResult {
    const results: EditResult[] = [];
    try {
      for (const c of this.commands) results.push(c.apply(workbook));
    } catch (error) {
      // A refused step: take back the ones before it, so the workbook is as it was.
      for (let i = results.length - 1; i >= 0; i--) this.commands[i]!.undo(workbook);
      throw error;
    }
    return combineResults(results);
  }

  undo(workbook: Workbook): EditResult {
    // Reverse order so each step sees the workbook in the state it left behind.
    return combineResults([...this.commands].reverse().map((c) => c.undo(workbook)));
  }
}

export function combineResults(results: readonly EditResult[]): EditResult {
  const cells: ChangedCell[] = [];
  const rows: ChangedLines[] = [];
  const columns: ChangedLines[] = [];
  let structural = false;
  for (const r of results) {
    cells.push(...r.changedCells);
    rows.push(...r.changedRows);
    columns.push(...r.changedColumns);
    structural ||= r.structural;
  }
  return result(cells, rows, columns, structural);
}

/** The named sheet, or throw: a command run against a workbook that lost its sheet is a bug. */
export function requireSheet(workbook: Workbook, name: string): Worksheet {
  const sheet = workbook.findSheet(name);
  if (!sheet) throw new Error(`Sheet '${name}' not found.`);
  return sheet;
}
