/** The parts of a keydown that tell a typed character from a composition. */
export interface KeyLike {
  readonly key: string;
  readonly keyCode: number;
  readonly isComposing: boolean;
}

/** keyCode browsers report for a key an IME (or a dead key, in some) is processing. */
const IME_KEY_CODE = 229;

/**
 * Whether a keydown on the grid starts a composed character: a dead key (´ then e for é) or an IME
 * (keyCode 229, `Process`). Its text only arrives later as input, and only in an editable element,
 * so the grid must open its editor now, without preventing the key, for the character to land.
 */
export function startsComposition(e: KeyLike): boolean {
  return e.key === 'Dead' || e.key === 'Process' || e.isComposing || e.keyCode === IME_KEY_CODE;
}

/**
 * A run of Tab presses along one row: where it started and where the last Tab landed. Enter at the
 * end of a run goes back to the column the run started in, one row down (Calc and Excel do this,
 * so a row of a form can be typed with Tab and the next row started with Enter).
 */
export interface TabRun {
  readonly row: number;
  /** The column the first Tab was pressed in. */
  readonly startColumn: number;
  /** Where the last Tab landed: the run only holds while the active cell is still there. */
  readonly column: number;
}

interface At {
  readonly row: number;
  readonly column: number;
}

/**
 * The run after a Tab (or Shift+Tab) from `from` to `to`; a Tab that changed rows ends it. A
 * Shift+Tab past the start moves the start with it, so Enter never steps back to the right.
 */
export function afterTab(run: TabRun | null, from: At, to: At): TabRun | null {
  if (to.row !== from.row) return null;
  const continues = run !== null && run.row === from.row && run.column === from.column;
  const start = continues ? run.startColumn : from.column;
  return { row: to.row, startColumn: Math.min(start, to.column), column: to.column };
}

/** Where Enter starts its step down from: the run's first column while the run holds, else `at`. */
export function enterOrigin(run: TabRun | null, at: At): At {
  if (run === null || run.row !== at.row || run.column !== at.column) return at;
  return { row: at.row, column: run.startColumn };
}
