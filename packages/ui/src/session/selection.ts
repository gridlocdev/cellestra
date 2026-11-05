import { signal, computed } from '@preact/signals';

export interface GridRange {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

export interface SelectionState {
  /** The active cell (where typing goes). */
  readonly row: number;
  readonly column: number;
  /** The fixed corner of the last range (shift-extend moves the other corner). */
  readonly anchorRow: number;
  readonly anchorColumn: number;
  /** All ranges (⌘-click adds); the last one contains the active cell. */
  readonly ranges: readonly GridRange[];
}

export function rangeOf(r0: number, c0: number, r1: number, c1: number): GridRange {
  return { top: Math.min(r0, r1), left: Math.min(c0, c1), bottom: Math.max(r0, r1), right: Math.max(c0, c1) };
}

export const MAX_ROW = 1_048_575;
export const MAX_COLUMN = 16_383;

/** The grid selection as a signal. Pure state; geometry and merges live in the grid. */
export class Selection {
  readonly state = signal<SelectionState>({
    row: 0,
    column: 0,
    anchorRow: 0,
    anchorColumn: 0,
    ranges: [rangeOf(0, 0, 0, 0)],
  });

  readonly active = computed(() => ({ row: this.state.value.row, column: this.state.value.column }));
  /** The last (primary) range. */
  readonly range = computed(() => this.state.value.ranges[this.state.value.ranges.length - 1]!);

  /** Move the active cell; collapses to a single range unless `extend`. */
  select(row: number, column: number, opts: { extend?: boolean; add?: boolean } = {}): void {
    const s = this.state.value;
    if (opts.extend) {
      const ranges = [...s.ranges.slice(0, -1), rangeOf(s.anchorRow, s.anchorColumn, row, column)];
      // Extending keeps the active cell at the anchor (like Calc/Excel).
      this.state.value = { ...s, ranges };
      return;
    }
    const r = rangeOf(row, column, row, column);
    this.state.value = {
      row,
      column,
      anchorRow: row,
      anchorColumn: column,
      ranges: opts.add ? [...s.ranges, r] : [r],
    };
  }

  /** Select a whole range with the active cell at its top-left (or a given cell inside it). */
  selectRange(range: GridRange, active?: { row: number; column: number }): void {
    const row = active?.row ?? range.top;
    const column = active?.column ?? range.left;
    this.state.value = { row, column, anchorRow: range.top, anchorColumn: range.left, ranges: [range] };
  }

  /** The moving corner of the last range (the one shift+arrow steps). */
  get extentCorner(): { row: number; column: number } {
    const s = this.state.value;
    const r = this.range.value;
    return {
      row: s.anchorRow === r.top ? r.bottom : r.top,
      column: s.anchorColumn === r.left ? r.right : r.left,
    };
  }

  /** Within a multi-cell range, Enter/Tab cycle the active cell instead of leaving the range. */
  cycle(rowStep: number, columnStep: number): boolean {
    const s = this.state.value;
    const r = this.range.value;
    if (r.top === r.bottom && r.left === r.right) return false;
    let { row, column } = s;
    if (columnStep !== 0) {
      column += columnStep;
      if (column > r.right) {
        column = r.left;
        row = row >= r.bottom ? r.top : row + 1;
      } else if (column < r.left) {
        column = r.right;
        row = row <= r.top ? r.bottom : row - 1;
      }
    } else {
      row += rowStep;
      if (row > r.bottom) {
        row = r.top;
        column = column >= r.right ? r.left : column + 1;
      } else if (row < r.top) {
        row = r.bottom;
        column = column <= r.left ? r.right : column - 1;
      }
    }
    this.state.value = { ...s, row, column };
    return true;
  }
}
