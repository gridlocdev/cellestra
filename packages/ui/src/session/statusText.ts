import { CellRange, type CalculationMode, type ChangedCell, type EditResult } from '@cellestra/core';
import { MAX_COLUMN, MAX_ROW, type GridRange } from './selection.ts';

/**
 * The status bar's words (PRD §6.5): what the last action did, when the file was saved, and the
 * calculation mode. Pure, so the session and the tests share them.
 */

/** Past tenses of the verbs command labels start with ("Insert Rows" → "Inserted rows"). */
const PAST_TENSE: Readonly<Record<string, string>> = {
  add: 'Added',
  apply: 'Applied',
  autofit: 'Autofitted',
  check: 'Checked',
  clean: 'Cleaned',
  clear: 'Cleared',
  create: 'Created',
  cut: 'Cut',
  decrease: 'Decreased',
  delete: 'Deleted',
  edit: 'Edited',
  fill: 'Filled',
  format: 'Formatted',
  freeze: 'Froze',
  hide: 'Hid',
  increase: 'Increased',
  insert: 'Inserted',
  merge: 'Merged',
  move: 'Moved',
  paste: 'Pasted',
  remove: 'Removed',
  rename: 'Renamed',
  reset: 'Reset',
  resize: 'Resized',
  set: 'Set',
  show: 'Showed',
  sort: 'Sorted',
  split: 'Split',
  transpose: 'Transposed',
  unfreeze: 'Unfroze',
  uncheck: 'Unchecked',
  unhide: 'Unhid',
  unmerge: 'Unmerged',
};

/**
 * A command label as something that has happened: "Paste" → "Pasted", "Insert Rows" → "Inserted
 * rows". A label that doesn't start with a known verb ("Bold", "Page Setup") is kept as it is.
 */
export function pastTense(label: string): { text: string; verb: boolean } {
  const space = label.indexOf(' ');
  const first = space < 0 ? label : label.slice(0, space);
  const past = PAST_TENSE[first.toLowerCase()];
  if (!past) return { text: label, verb: false };
  return { text: space < 0 ? past : past + label.slice(space).toLowerCase(), verb: true };
}

/**
 * Where an edit landed, in a few characters: one cell's address, the range when the changed cells
 * fill a rectangle on one sheet, otherwise how many cells changed.
 */
export function describeCells(cells: readonly ChangedCell[]): string | null {
  const first = cells[0];
  if (!first) return null;
  // A cell can be reported twice (a value and its recalculated formula), so count positions.
  const seen = new Set<string>();
  let { row: top, row: bottom, column: left, column: right } = first;
  let oneSheet = true;
  for (const c of cells) {
    if (c.sheet !== first.sheet) oneSheet = false;
    seen.add(`${c.sheet}\u0000${c.row}\u0000${c.column}`);
    top = Math.min(top, c.row);
    bottom = Math.max(bottom, c.row);
    left = Math.min(left, c.column);
    right = Math.max(right, c.column);
  }
  const area = (bottom - top + 1) * (right - left + 1);
  if (oneSheet && seen.size === area) return describeRange({ top, left, bottom, right });
  return `${seen.size.toLocaleString('en-US')} ${seen.size === 1 ? 'cell' : 'cells'}`;
}

/** A range in A1 form: "B2" or "A1:C20", or "B:C" and "2:4" for whole columns and rows. */
export function describeRange(r: GridRange): string {
  const a1 = CellRange.fromBounds(r.top, r.left, r.bottom, r.right).toString();
  const wholeColumns = r.top === 0 && r.bottom >= MAX_ROW;
  const wholeRows = r.left === 0 && r.right >= MAX_COLUMN;
  if (wholeColumns === wholeRows) return a1;
  const [from, to = from] = a1.split(':') as [string, string?];
  const part = (ref: string) => (wholeColumns ? ref.replace(/\d+$/, '') : ref.replace(/^[A-Z]+/, ''));
  return `${part(from)}:${part(to)}`;
}

/**
 * The last-action line for an edit: "Pasted A1:C4", "Sorted 40 cells", "Inserted rows",
 * "Bold · B2:C4". `selection` stands in for the target when the edit reports no cells (a style).
 */
export function narrateEdit(label: string, result: EditResult, selection: GridRange | null): string {
  const { text, verb } = pastTense(label);
  if (result.structural) return text;
  const target = describeCells(result.changedCells) ?? (selection === null ? null : describeRange(selection));
  if (target === null) return text;
  return verb ? `${text} ${target}` : `${text} · ${target}`;
}

/** "Undid Paste" / "Redid Insert Rows": undo and redo name the step they moved over. */
export function narrateHistory(kind: 'undo' | 'redo', label: string): string {
  return `${kind === 'undo' ? 'Undid' : 'Redid'} ${label}`;
}

/** Hours and minutes on the local clock: a status line for a person, not a cell value. */
export function clockTime(at: Date): string {
  return `${String(at.getHours()).padStart(2, '0')}:${String(at.getMinutes()).padStart(2, '0')}`;
}

/**
 * The save-state indicator: "Saved 14:05" while the document matches its last save, "Edited" once
 * it doesn't, "Not saved" for a document that has never had a file. Empty for a fresh blank one.
 */
export function saveStateText(modified: boolean, hasFile: boolean, savedAt: Date | null): string {
  if (modified) return hasFile ? 'Edited' : 'Not saved';
  return savedAt ? `Saved ${clockTime(savedAt)}` : '';
}

/** The calculation-mode badge. */
export function calculationBadge(mode: CalculationMode): string {
  return mode === 'manual' ? 'Manual' : 'Auto';
}
