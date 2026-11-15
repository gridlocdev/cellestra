import type { GridRange } from '../session/selection.ts';

/**
 * How the active cell moves on the sheet, apart from drawing and the DOM: which cell an arrow,
 * Enter, Tab or Page key lands on, past hidden lines and merged cells, and what a key on the grid
 * (not in the cell editor) asks for.
 */

interface At {
  readonly row: number;
  readonly column: number;
}

/** What stepping needs from a line axis: how to count visible lines, and which are hidden. */
export interface StepAxis {
  step(index: number, delta: number): number;
  isHidden(index: number): boolean;
}

export interface StepAxes {
  readonly rows: StepAxis;
  readonly columns: StepAxis;
}

/** What stepping needs from the sheet: one step that treats a merged cell as one cell. */
export interface StepSheet {
  stepOverMerges(row: number, column: number, rowDelta: number, columnDelta: number): [number, number];
}

/**
 * How a step counts: `cell` moves one cell (a merge is one cell), `extend` moves the far corner of
 * a selection one visible line, `page` moves the given number of visible rows.
 */
export type StepKind = 'cell' | 'extend' | 'page';

/** Where a step of `dr` rows and `dc` columns from `from` lands. Never on a hidden line. */
export function stepFrom(
  from: At,
  dr: number,
  dc: number,
  kind: StepKind,
  axes: StepAxes,
  sheet: StepSheet,
): At {
  if (kind === 'extend') {
    return {
      row: dr ? axes.rows.step(from.row, dr) : from.row,
      column: dc ? axes.columns.step(from.column, dc) : from.column,
    };
  }
  if (kind === 'page') {
    // Count visible rows, so hidden ones don't shorten (or overshoot) a page.
    const row = axes.rows.step(from.row, dr);
    return { row: axes.rows.isHidden(row) ? from.row : row, column: from.column };
  }
  let [row, column] = sheet.stepOverMerges(from.row, from.column, dr, dc);
  // Past hidden lines; with nothing visible that way, stay put rather than land on a hidden one.
  if (dr && axes.rows.isHidden(row)) row = axes.rows.step(row, dr);
  if (dr && axes.rows.isHidden(row)) row = from.row;
  if (dc && axes.columns.isHidden(column)) column = axes.columns.step(column, dc);
  if (dc && axes.columns.isHidden(column)) column = from.column;
  return { row, column };
}

/** The first visible line: 0, unless it is hidden. */
export function firstVisible(axis: StepAxis): number {
  return axis.isHidden(0) ? axis.step(0, 1) : 0;
}

/** Widen a range until no merge straddles its edge (merges can chain, so to a fixed point). */
export function growOverMerges(merges: readonly { readonly range: GridRange }[], r: GridRange): GridRange {
  if (merges.length === 0) return r;
  let { top, left, bottom, right } = r;
  for (let changed = true; changed;) {
    changed = false;
    for (const m of merges) {
      const g = m.range;
      if (g.bottom < top || g.top > bottom || g.right < left || g.left > right) continue;
      if (g.top < top || g.left < left || g.bottom > bottom || g.right > right) {
        top = Math.min(top, g.top);
        left = Math.min(left, g.left);
        bottom = Math.max(bottom, g.bottom);
        right = Math.max(right, g.right);
        changed = true;
      }
    }
  }
  return { top, left, bottom, right };
}

/** The parts of a keydown the grid reads. */
export interface GridKey {
  readonly key: string;
  readonly shiftKey: boolean;
  readonly altKey: boolean;
  readonly metaKey: boolean;
  readonly ctrlKey: boolean;
}

/** What a key pressed on the grid asks for. */
export type GridKeyAction =
  | {
      readonly kind: 'move';
      readonly rows: number;
      readonly columns: number;
      /** Shift: stretch the selection rather than move the cell. */
      readonly extend: boolean;
      /** Enter and Tab: move within a selected range, and remember a run of Tabs. */
      readonly cycle: boolean;
      /** ⌘/Ctrl: jump to the edge of the data. */
      readonly jump: boolean;
    }
  | { readonly kind: 'page'; readonly down: boolean; readonly extend: boolean }
  | { readonly kind: 'home'; readonly top: boolean; readonly extend: boolean }
  | { readonly kind: 'edit' }
  | { readonly kind: 'clear' }
  | { readonly kind: 'escape' }
  | { readonly kind: 'type'; readonly text: string };

const ARROWS: Readonly<Record<string, readonly [number, number]>> = {
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1],
};

/** The action a key on the grid stands for, or null for a key the grid leaves alone. */
export function gridKeyAction(e: GridKey): GridKeyAction | null {
  const mod = e.metaKey || e.ctrlKey;
  const arrow = ARROWS[e.key];
  if (arrow) {
    return { kind: 'move', rows: arrow[0], columns: arrow[1], extend: e.shiftKey, cycle: false, jump: mod };
  }
  switch (e.key) {
    case 'Enter':
      return { kind: 'move', rows: e.shiftKey ? -1 : 1, columns: 0, extend: false, cycle: true, jump: false };
    case 'Tab':
      return { kind: 'move', rows: 0, columns: e.shiftKey ? -1 : 1, extend: false, cycle: true, jump: false };
    case 'PageDown':
    case 'PageUp':
      return { kind: 'page', down: e.key === 'PageDown', extend: e.shiftKey };
    case 'Home':
      return { kind: 'home', top: mod, extend: e.shiftKey };
    case 'F2':
      return { kind: 'edit' };
    case 'Delete':
    case 'Backspace':
      return { kind: 'clear' };
    case 'Escape':
      return { kind: 'escape' };
  }
  // Typing a printable character replaces the cell's content and starts editing. ⌥ (macOS) and
  // AltGr (Ctrl+Alt elsewhere) type characters too; ⌘ and plain Ctrl are shortcuts.
  if (e.key.length === 1 && !e.metaKey && (!e.ctrlKey || e.altKey)) return { kind: 'type', text: e.key };
  return null;
}
