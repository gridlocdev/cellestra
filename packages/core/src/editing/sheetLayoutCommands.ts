import type { DataValidation, FilterSpec, SortSpec, Table } from '../data/types.ts';
import type { CellRange } from '../model/cellRange.ts';
import type { MergedRegion } from '../model/mergedRegion.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
import { StylePool } from '../styling/stylePool.ts';
import { restoreLines, type LineState } from './dataCommandSupport.ts';
import { ActionCommand, EditResult, requireSheet, type EditCommand } from './editCommand.ts';

/** Undoable resets of a sheet's or range's structure: tables, merges, validation, filters, line styles. */
export const SheetLayoutCommands = {
  /** Reset every line style/hidden flag, filter/sort, table, validation, merge, and frozen pane. */
  resetSheetLayout(sheet: string): EditCommand {
    let rowState: LineState[] = [];
    let columnState: LineState[] = [];
    let filter: CellRange | null = null;
    let filterSpec: FilterSpec | null = null;
    let sortState: SortSpec | null = null;
    let tables: Table[] = [];
    let merges: MergedRegion[] = [];
    let validations: DataValidation[] = [];
    let frozenRows = 0;
    let frozenColumns = 0;
    const everyColumn = EditResult.columns(sheet, 0, SheetDimensions.maxColumns - 1);
    return new ActionCommand(
      'Reset Sheet Data',
      (wb) => {
        const ws = requireSheet(wb, sheet);
        rowState = [...ws.rows].map(([k, r]) => [k, r.styleId, r.manuallyHidden, r.filtered]);
        columnState = [...ws.columns].map(([k, c]) => [k, c.styleId, c.hidden, false]);
        filter = ws.activeFilter;
        filterSpec = ws.activeFilterSpec;
        sortState = ws.activeSort;
        tables = [...ws.tables];
        merges = [...ws.merges];
        validations = [...ws.validations];
        frozenRows = ws.frozenRows;
        frozenColumns = ws.frozenColumns;
        for (const row of ws.rows.values()) {
          row.styleId = StylePool.defaultId;
          row.manuallyHidden = false;
          row.filtered = false;
        }
        for (const col of ws.columns.values()) {
          col.styleId = StylePool.defaultId;
          col.hidden = false;
        }
        ws.activeFilter = null;
        ws.activeFilterSpec = null;
        ws.activeSort = null;
        for (const t of tables) ws.removeTable(t);
        for (const v of validations) ws.removeValidation(v);
        ws.replaceMerges([]);
        ws.frozenRows = 0;
        ws.frozenColumns = 0;
        return everyColumn;
      },
      (wb) => {
        const ws = requireSheet(wb, sheet);
        restoreLines(ws, rowState, columnState);
        ws.activeFilter = filter;
        ws.activeFilterSpec = filterSpec;
        ws.activeSort = sortState;
        for (const t of tables) ws.addTable(t);
        for (const v of validations) ws.addValidation(v);
        ws.replaceMerges(merges);
        ws.frozenRows = frozenRows;
        ws.frozenColumns = frozenColumns;
        return everyColumn;
      },
    );
  },
} as const;
