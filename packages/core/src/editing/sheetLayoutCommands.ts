import type { DataValidation, FilterSpec, SortSpec, Table } from '../data/types.ts';
import type { CellRange } from '../model/cellRange.ts';
import type { MergedRegion } from '../model/mergedRegion.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
import { StylePool } from '../styling/stylePool.ts';
import { restoreLines, type LineState } from './dataCommandSupport.ts';
import { ActionCommand, combineResults, EditResult, requireSheet, type EditCommand } from './editCommand.ts';
import { coversWholeColumns } from './editCommands.ts';

/** Undoable resets of a sheet's or range's structure: tables, merges, validation, filters, line styles. */
export const SheetLayoutCommands = {
  /**
   * The structural half of Clear All: tables, merges, and validations touching the range go, an
   * intersecting AutoFilter is dropped (its rows unhidden), and whole rows/columns lose their line
   * style and hidden flag.
   */
  clearRangeStructure(sheet: string, range: CellRange): EditCommand {
    let removedTables: Table[] = [];
    let removedMerges: MergedRegion[] = [];
    let removedValidations: DataValidation[] = [];
    let clearedFilter = false;
    let filter: CellRange | null = null;
    let filterSpec: FilterSpec | null = null;
    let sortState: SortSpec | null = null;
    let rowState: LineState[] = [];
    let columnState: LineState[] = [];
    /** The range's rows and the rows of the merges it removed; its columns when they were whole. */
    const changed = (): EditResult =>
      combineResults([
        EditResult.rows(sheet, range.top, range.bottom),
        ...removedMerges.map((m) => EditResult.rows(sheet, m.range.top, m.range.bottom)),
        coversWholeColumns(range) ? EditResult.columns(sheet, range.left, range.right) : EditResult.none,
      ]);
    return new ActionCommand(
      'Clear All',
      (wb) => {
        const ws = requireSheet(wb, sheet);
        removedTables = ws.tables.filter((t) => t.range.intersects(range));
        for (const t of removedTables) ws.removeTable(t);
        removedMerges = ws.merges.filter((m) => m.range.intersects(range));
        if (removedMerges.length > 0) ws.replaceMerges(ws.merges.filter((m) => !m.range.intersects(range)));
        removedValidations = ws.validations.filter((v) => v.target.intersects(range));
        for (const v of removedValidations) ws.removeValidation(v);

        const active = ws.activeFilter;
        clearedFilter = active !== null && active.intersects(range);
        const filterRange = clearedFilter ? active : null;
        if (clearedFilter) {
          filter = ws.activeFilter;
          filterSpec = ws.activeFilterSpec;
          sortState = ws.activeSort;
        }
        const wholeColumns = range.top === 0 && range.bottom >= SheetDimensions.maxRows - 1;
        const wholeRows = range.left === 0 && range.right >= SheetDimensions.maxColumns - 1;
        const touchedRows = new Set<number>();
        if (wholeRows)
          for (const r of ws.rows.keys()) if (r >= range.top && r <= range.bottom) touchedRows.add(r);
        if (filterRange) {
          for (const r of ws.rows.keys()) {
            if (r >= filterRange.top && r <= filterRange.bottom && ws.isRowFiltered(r)) touchedRows.add(r);
          }
        }
        const sortedRows = [...touchedRows].sort((a, b) => a - b);
        const touchedColumns = wholeColumns
          ? [...ws.columns.keys()].filter((c) => c >= range.left && c <= range.right)
          : [];
        rowState = sortedRows.map((r) => {
          const row = ws.getOrCreateRow(r);
          return [r, row.styleId, row.manuallyHidden, row.filtered];
        });
        columnState = touchedColumns.map((c) => {
          const col = ws.getOrCreateColumn(c);
          return [c, col.styleId, col.hidden, false];
        });
        if (clearedFilter) {
          ws.activeFilter = null;
          ws.activeFilterSpec = null;
          ws.activeSort = null;
        }
        // Whole rows lose their layout; a cleared filter shows only the rows it hid.
        for (const r of sortedRows) {
          const row = ws.getOrCreateRow(r);
          const whole = wholeRows && r >= range.top && r <= range.bottom;
          if (whole) {
            row.styleId = StylePool.defaultId;
            row.manuallyHidden = false;
          }
          if (whole || clearedFilter) row.filtered = false;
        }
        for (const c of touchedColumns) {
          const col = ws.getOrCreateColumn(c);
          col.styleId = StylePool.defaultId;
          col.hidden = false;
        }
        return changed();
      },
      (wb) => {
        const ws = requireSheet(wb, sheet);
        for (const t of removedTables) ws.addTable(t);
        if (removedMerges.length > 0) ws.replaceMerges([...ws.merges, ...removedMerges]);
        for (const v of removedValidations) ws.addValidation(v);
        if (clearedFilter) {
          ws.activeFilter = filter;
          ws.activeFilterSpec = filterSpec;
          ws.activeSort = sortState;
        }
        restoreLines(ws, rowState, columnState);
        return changed();
      },
    );
  },

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
