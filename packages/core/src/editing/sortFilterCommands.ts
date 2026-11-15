import { FilterEngine } from '../data/filterEngine.ts';
import { SortEngine } from '../data/sortEngine.ts';
import type { FilterSpec, SortSpec, Table } from '../data/types.ts';
import type { CellRange } from '../model/cellRange.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { CellSnapshotCommand } from './cellSnapshotCommand.ts';
import { filteredRowsCommand, inDocumentLocale, rowVisibilityCommand } from './dataCommandSupport.ts';
import {
  ActionCommand,
  CompositeEditCommand,
  EditResult,
  requireSheet,
  type EditCommand,
} from './editCommand.ts';
import { clampToUsed, EditCommands } from './editCommands.ts';
import { TableCommands } from './tableCommands.ts';

/** Undoable sort and filter edits, for plain ranges, the sheet's AutoFilter and tables. */
export const SortFilterCommands = {
  /** Sort the range (clamped to the used range) as one undo step. */
  /** Text collates in `locale`, else in the document locale. */
  sort: (sheet: string, range: CellRange, spec: SortSpec, locale?: string): EditCommand =>
    inDocumentLocale(
      'Sort',
      locale,
      (loc) =>
        // Sorting re-applies a filter over the range, which changes which rows are hidden; the first
        // step keeps those flags so undo shows the same rows again.
        new CompositeEditCommand('Sort', [
          filteredRowsCommand(sheet, range),
          CellSnapshotCommand.overRange(
            'Sort',
            sheet,
            (ws) => clampToUsed(ws, range),
            (ws) => SortEngine.sort(ws, clampToUsed(ws, range) ?? range, spec, loc),
          ),
        ]),
    ),

  sortByColumn: (
    sheet: string,
    range: CellRange,
    column: number,
    ascending: boolean,
    hasHeader = true,
    locale?: string,
  ): EditCommand =>
    SortFilterCommands.sort(sheet, range, { levels: [{ column, ascending }], hasHeader }, locale),

  applyFilter: (sheet: string, range: CellRange, spec: FilterSpec, locale: string): EditCommand =>
    rowVisibilityCommand('Filter', sheet, range, (ws) => FilterEngine.apply(ws, range, spec, locale)),

  clearFilter: (sheet: string, range: CellRange): EditCommand =>
    rowVisibilityCommand('Clear Filter', sheet, range, (ws) => FilterEngine.clear(ws, range)),

  /**
   * The Sort & Filter dialog's single undo step for a plain range: records the sheet AutoFilter
   * state, applies (or clears) the filter, then sorts — undo restores order, visibility, and state.
   */
  applySortAndFilter(
    sheet: string,
    range: CellRange,
    sort: SortSpec | null,
    filter: FilterSpec | null,
    locale?: string,
  ): EditCommand {
    return inDocumentLocale('Sort & Filter', locale, (loc) => {
      const steps: EditCommand[] = [
        setSheetFilterState(sheet, sort || filter ? range : null, filter, sort),
        filter
          ? SortFilterCommands.applyFilter(sheet, range, filter, loc)
          : SortFilterCommands.clearFilter(sheet, range),
      ];
      if (sort && sort.levels.length > 0) steps.push(SortFilterCommands.sort(sheet, range, sort, loc));
      return EditCommands.composite('Sort & Filter', steps);
    });
  },

  /** As `applySortAndFilter`, but the state lives on the table record, not the sheet AutoFilter. */
  applyTableSortAndFilter(
    sheet: string,
    table: Table,
    sort: SortSpec | null,
    filter: FilterSpec | null,
    locale?: string,
  ): EditCommand {
    return inDocumentLocale('Sort & Filter', locale, (loc) => {
      const range = table.range;
      const steps: EditCommand[] = [
        TableCommands.replaceTable(sheet, table, { ...table, filter, sort }),
        filter
          ? SortFilterCommands.applyFilter(sheet, range, filter, loc)
          : SortFilterCommands.clearFilter(sheet, range),
      ];
      if (sort && sort.levels.length > 0) steps.push(SortFilterCommands.sort(sheet, range, sort, loc));
      return EditCommands.composite('Sort & Filter', steps);
    });
  },

  /** Undoable switch of the sheet's AutoFilter range (keeps its filter and sort only when the range stays). */
  setAutoFilter(sheet: string, range: CellRange | null): EditCommand {
    let before: [CellRange | null, FilterSpec | null, SortSpec | null] = [null, null, null];
    return new ActionCommand(
      range ? 'AutoFilter' : 'Remove AutoFilter',
      (wb) => {
        const ws = requireSheet(wb, sheet);
        before = [ws.activeFilter, ws.activeFilterSpec, ws.activeSort];
        const same = range && ws.activeFilter?.equals(range);
        ws.activeFilter = range;
        if (!same) {
          ws.activeFilterSpec = null;
          ws.activeSort = null;
        }
        return EditResult.none;
      },
      (wb) => {
        const ws = requireSheet(wb, sheet);
        [ws.activeFilter, ws.activeFilterSpec, ws.activeSort] = before;
        return EditResult.none;
      },
    );
  },

  /** Turn the sheet's AutoFilter off: its rows show again and its filter and sort go, in one step. */
  removeAutoFilter(worksheet: Worksheet, locale?: string): EditCommand | null {
    const range = worksheet.activeFilter;
    if (!range) return null;
    return EditCommands.composite('Remove AutoFilter', [
      SortFilterCommands.applySortAndFilter(worksheet.name, range, null, null, locale),
      SortFilterCommands.setAutoFilter(worksheet.name, null),
    ]);
  },

  /**
   * Sort & Filter over a table or a plain range, as one undo step. A plain AutoFilter left with no
   * filter or sort would vanish, so it is put back in the same step (keeping its buttons).
   */
  sortAndFilter(
    worksheet: Worksheet,
    target: { readonly range: CellRange; readonly table: Table | null },
    sort: SortSpec | null,
    filter: FilterSpec | null,
    locale?: string,
  ): EditCommand {
    const sheet = worksheet.name;
    const auto = worksheet.activeFilter;
    const steps: EditCommand[] = [];
    if (target.table) {
      if (auto && auto.equals(target.table.range))
        steps.push(SortFilterCommands.applySortAndFilter(sheet, auto, null, null, locale));
      steps.push(SortFilterCommands.applyTableSortAndFilter(sheet, target.table, sort, filter, locale));
    } else {
      steps.push(SortFilterCommands.applySortAndFilter(sheet, target.range, sort, filter, locale));
      if (!filter && !sort && auto?.equals(target.range))
        steps.push(SortFilterCommands.setAutoFilter(sheet, target.range));
    }
    return steps.length === 1 ? steps[0]! : new CompositeEditCommand('Sort & Filter', steps);
  },
} as const;

function setSheetFilterState(
  sheet: string,
  filterRange: CellRange | null,
  filterSpec: FilterSpec | null,
  sort: SortSpec | null,
): EditCommand {
  let oldFilter: CellRange | null = null;
  let oldSpec: FilterSpec | null = null;
  let oldSort: SortSpec | null = null;
  return new ActionCommand(
    'Filter',
    (wb) => {
      const ws = requireSheet(wb, sheet);
      oldFilter = ws.activeFilter;
      oldSpec = ws.activeFilterSpec;
      oldSort = ws.activeSort;
      ws.activeFilter = filterRange;
      ws.activeFilterSpec = filterSpec;
      ws.activeSort = sort;
      return EditResult.none;
    },
    (wb) => {
      const ws = requireSheet(wb, sheet);
      ws.activeFilter = oldFilter;
      ws.activeFilterSpec = oldSpec;
      ws.activeSort = oldSort;
      return EditResult.none;
    },
  );
}
