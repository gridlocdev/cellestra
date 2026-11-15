import type { CellRange } from '../model/cellRange.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { ActionCommand, EditResult, requireSheet, type EditCommand } from './editCommand.ts';

/** Shared plumbing for the data commands (sort, filter, tables, layout). */

/** A row's or column's style and visibility; `filtered` is always false for a column. */
export type LineState = [index: number, styleId: number, hidden: boolean, filtered: boolean];

/**
 * The command `build` makes for `locale`, or, without one, for the workbook's document locale (its
 * effective locale) as it is when the command first applies.
 */
export function inDocumentLocale(
  label: string,
  locale: string | undefined,
  build: (locale: string) => EditCommand,
): EditCommand {
  if (locale !== undefined) return build(locale);
  let inner: EditCommand | null = null;
  return new ActionCommand(
    label,
    (wb) => {
      inner ??= build(wb.settings.effectiveLocale);
      return inner.apply(wb);
    },
    (wb) => inner?.undo(wb) ?? EditResult.none,
  );
}

/**
 * Changes nothing; keeps which rows of every filter that overlaps `range` (the sheet's filter and
 * table filters) are hidden, and puts exactly those flags back on undo.
 */
export function filteredRowsCommand(sheet: string, range: CellRange): EditCommand {
  let filtered: CellRange[] = [];
  let hiddenBefore = new Set<number>();
  return new ActionCommand(
    'Sort',
    (wb) => {
      const ws = requireSheet(wb, sheet);
      filtered = [];
      if (ws.activeFilter && ws.activeFilterSpec && ws.activeFilter.intersects(range))
        filtered.push(ws.activeFilter);
      for (const t of ws.tables) if (t.filter && t.range.intersects(range)) filtered.push(t.range);
      hiddenBefore = new Set();
      for (const r of filtered) for (const row of filteredRowsIn(ws, r)) hiddenBefore.add(row);
      return EditResult.none;
    },
    (wb) => {
      const ws = requireSheet(wb, sheet);
      for (const r of filtered) {
        for (const [row, model] of ws.rows) {
          if (row >= r.top && row <= r.bottom) model.filtered = hiddenBefore.has(row);
        }
      }
      for (const row of hiddenBefore) if (!ws.isRowFiltered(row)) ws.getOrCreateRow(row).filtered = true;
      return EditResult.none;
    },
  );
}

/**
 * Run a change to which rows a filter hides over the range; undo puts back the filtered flag of
 * each row it changed (found from the sparse row models, so a whole-column range costs nothing
 * extra). Rows hidden by hand are not the filter's and are left alone.
 */
export function rowVisibilityCommand(
  label: string,
  sheet: string,
  range: CellRange,
  apply: (ws: Worksheet) => void,
): EditCommand {
  let changed: (readonly [row: number, filtered: boolean])[] = [];
  return new ActionCommand(
    label,
    (wb) => {
      const ws = requireSheet(wb, sheet);
      const hiddenBefore = filteredRowsIn(ws, range);
      apply(ws);
      const hiddenAfter = filteredRowsIn(ws, range);
      changed = [];
      for (const row of hiddenBefore) if (!hiddenAfter.has(row)) changed.push([row, true]);
      for (const row of hiddenAfter) if (!hiddenBefore.has(row)) changed.push([row, false]);
      return changedRowsResult(sheet, changed);
    },
    (wb) => {
      const ws = requireSheet(wb, sheet);
      for (const [row, filtered] of changed) ws.getOrCreateRow(row).filtered = filtered;
      return changedRowsResult(sheet, changed);
    },
  );
}

/** The span of rows whose hidden flag changed, so the session re-measures them and recalculates SUBTOTAL. */
function changedRowsResult(
  sheet: string,
  changed: readonly (readonly [row: number, hidden: boolean])[],
): EditResult {
  if (changed.length === 0) return EditResult.none;
  let first = changed[0]![0];
  let last = first;
  for (const [row] of changed) {
    if (row < first) first = row;
    if (row > last) last = row;
  }
  return EditResult.rows(sheet, first, last);
}

function filteredRowsIn(ws: Worksheet, range: CellRange): Set<number> {
  const out = new Set<number>();
  for (const [row, model] of ws.rows)
    if (model.filtered && row >= range.top && row <= range.bottom) out.add(row);
  return out;
}

export function restoreLines(ws: Worksheet, rows: readonly LineState[], columns: readonly LineState[]): void {
  for (const [k, styleId, hidden, filtered] of rows) {
    const row = ws.getOrCreateRow(k);
    row.styleId = styleId;
    row.manuallyHidden = hidden;
    row.filtered = filtered;
  }
  for (const [k, styleId, hidden] of columns) {
    const col = ws.getOrCreateColumn(k);
    col.styleId = styleId;
    col.hidden = hidden;
  }
}
