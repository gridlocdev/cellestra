import { DuplicateRemover, TextToColumns } from '../data/dataTools.ts';
import { FilterEngine } from '../data/filterEngine.ts';
import { SortEngine } from '../data/sortEngine.ts';
import { TableAutoExpand } from '../data/tables.ts';
import type { DataValidation, FilterSpec, SortSpec, Table } from '../data/types.ts';
import { CellRange } from '../model/cellRange.ts';
import type { MergedRegion } from '../model/mergedRegion.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import type { CellStyle } from '../styling/cellStyle.ts';
import { StylePool } from '../styling/stylePool.ts';
import { CellValue } from '../values/cellValue.ts';
import { CellSnapshotCommand } from './cellSnapshotCommand.ts';
import {
  ActionCommand,
  combineResults,
  CompositeEditCommand,
  EditResult,
  requireSheet,
  type EditCommand,
} from './editCommand.ts';
import { clampToUsed, coversWholeColumns, EditCommands } from './editCommands.ts';

/** A row's or column's style and visibility; `filtered` is always false for a column. */
type LineState = [index: number, styleId: number, hidden: boolean, filtered: boolean];

/** Undoable data-menu edits: sort, filter, duplicates, text to columns, tables, validation. */
export const DataEditCommands = {
  // ---- Sort -------------------------------------------------------------

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
    DataEditCommands.sort(sheet, range, { levels: [{ column, ascending }], hasHeader }, locale),

  // ---- Remove duplicates ------------------------------------------------

  /** The command, plus a getter for how many rows its last apply removed. */
  removeDuplicates(
    sheet: string,
    range: CellRange,
    keyColumns: readonly number[],
    hasHeader: boolean,
  ): { readonly command: EditCommand; readonly removedCount: () => number } {
    let removed = 0;
    // Cut off below and right of the used range, so a whole-column selection doesn't snapshot a
    // million rows. The top row stays: it may be the header.
    const trimmed = (ws: Worksheet): CellRange | null => {
      const used = ws.getUsedRange();
      if (!used) return null;
      const bottom = Math.min(range.bottom, used.bottom);
      const right = Math.min(range.right, used.right);
      return bottom < range.top || right < range.left
        ? null
        : CellRange.fromBounds(range.top, range.left, bottom, right);
    };
    const command = new CellSnapshotCommand(
      'Remove Duplicates',
      sheet,
      (ws) => [...(trimmed(ws)?.positions() ?? [])],
      (ws) => {
        const r = trimmed(ws);
        removed = r ? DuplicateRemover.remove(ws, r, keyColumns, hasHeader) : 0;
      },
    );
    return { command, removedCount: () => removed };
  },

  // ---- Text to columns --------------------------------------------------

  /**
   * Split a column's text; snapshots a block as wide as the widest split, so it fully undoes. Only
   * the used rows of a whole-column selection are read. Fields are read as typed input in `locale`
   * (numbers, dates, TRUE/FALSE; by default the document locale), and a date gets its date format.
   */
  textToColumns(
    worksheet: Worksheet,
    sourceColumn: CellRange,
    delimiter: string,
    locale?: string,
  ): EditCommand {
    const source = (ws: Worksheet): CellRange | null => {
      const bottom = Math.min(sourceColumn.bottom, ws.cellExtent()?.lastRow ?? -1);
      return bottom < sourceColumn.top
        ? null
        : CellRange.fromBounds(sourceColumn.top, sourceColumn.left, bottom, sourceColumn.left);
    };
    const block = (ws: Worksheet): CellRange | null => {
      const src = source(ws);
      if (!src) return null;
      let maxFields = 1;
      for (let row = src.top; row <= src.bottom; row++) {
        const v = ws.getValue(row, src.left);
        if (v.kind === 'text')
          maxFields = Math.max(maxFields, TextToColumns.fields(v.value, delimiter).length);
      }
      const right = Math.min(src.left + maxFields - 1, SheetDimensions.maxColumns - 1);
      return CellRange.fromBounds(src.top, src.left, src.bottom, right);
    };
    let styles: StylePool | null = null;
    const split = new CellSnapshotCommand(
      'Text to Columns',
      worksheet.name,
      (ws) => [...(block(ws)?.positions() ?? [])],
      (ws) => {
        const src = source(ws);
        if (src) TextToColumns.split(ws, src, delimiter, { locale: documentLocale, styles });
      },
    );
    let documentLocale = locale;
    return new ActionCommand(
      'Text to Columns',
      (wb) => {
        styles = wb.styles;
        documentLocale = locale ?? wb.settings.effectiveLocale;
        return split.apply(wb);
      },
      (wb) => split.undo(wb),
    );
  },

  // ---- Filter -----------------------------------------------------------

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
          ? DataEditCommands.applyFilter(sheet, range, filter, loc)
          : DataEditCommands.clearFilter(sheet, range),
      ];
      if (sort && sort.levels.length > 0) steps.push(DataEditCommands.sort(sheet, range, sort, loc));
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
        DataEditCommands.replaceTable(sheet, table, { ...table, filter, sort }),
        filter
          ? DataEditCommands.applyFilter(sheet, range, filter, loc)
          : DataEditCommands.clearFilter(sheet, range),
      ];
      if (sort && sort.levels.length > 0) steps.push(DataEditCommands.sort(sheet, range, sort, loc));
      return EditCommands.composite('Sort & Filter', steps);
    });
  },

  // ---- Clear structure --------------------------------------------------

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

  // ---- Tables -----------------------------------------------------------

  createTable: (sheet: string, table: Table): EditCommand =>
    new ActionCommand(
      'Create Table',
      (wb) => {
        requireSheet(wb, sheet).addTable(table);
        return EditResult.none;
      },
      (wb) => {
        requireSheet(wb, sheet).removeTable(table);
        return EditResult.none;
      },
    ),

  /**
   * Create a table and, when it has a header, name its empty header cells ("Column 1", …) and bold
   * the header row — all one undo step.
   */
  createStyledTable(workbook: Workbook, sheet: string, table: Table): EditCommand {
    const add = DataEditCommands.createTable(sheet, table);
    if (!table.hasHeader) return add;
    const r = table.range;
    const headerRange = CellRange.fromBounds(r.top, r.left, r.top, r.right);
    const fillHeaders = new CellSnapshotCommand('Create Table', sheet, [...headerRange.positions()], (ws) => {
      for (let column = r.left; column <= r.right; column++) {
        if (ws.getValue(r.top, column).kind === 'empty') {
          ws.setValue(r.top, column, CellValue.text(`Column ${column - r.left + 1}`));
        }
      }
    });
    const style = EditCommands.applyStyle(workbook, sheet, headerRange, styleTableHeader, 'Create Table');
    return EditCommands.composite('Create Table', [add, fillHeaders, style]);
  },

  styleTableHeader: (style: CellStyle): CellStyle => styleTableHeader(style),
  clearTableHeader: (style: CellStyle): CellStyle => ({ ...style, font: { ...style.font, bold: false } }),

  /** Flip the table's header column and bold (or unbold) its first column below the header row. */
  setTableHeaderColumn(workbook: Workbook, sheet: string, table: Table, enabled: boolean): EditCommand {
    const replaced: Table = { ...table, hasHeaderColumn: enabled };
    const flip = swapTable('Header Column', sheet, table, replaced);
    const r = table.range;
    const top = r.top + (table.hasHeader ? 1 : 0);
    if (top > r.bottom) return flip; // header-only table: nothing to restyle
    const columnRange = CellRange.fromBounds(top, r.left, r.bottom, r.left);
    const style = EditCommands.applyStyle(
      workbook,
      sheet,
      columnRange,
      enabled ? styleTableHeader : DataEditCommands.clearTableHeader,
      'Header Column',
    );
    return EditCommands.composite('Header Column', [flip, style]);
  },

  setTableBandedRows: (sheet: string, table: Table, enabled: boolean): EditCommand =>
    swapTable('Banded Rows', sheet, table, { ...table, hasBandedRows: enabled }),

  removeTable: (sheet: string, table: Table): EditCommand =>
    new ActionCommand(
      'Remove Table',
      (wb) => {
        requireSheet(wb, sheet).removeTable(table);
        return EditResult.none;
      },
      (wb) => {
        requireSheet(wb, sheet).addTable(table);
        return EditResult.none;
      },
    ),

  replaceTable: (sheet: string, existing: Table, replacement: Table): EditCommand =>
    swapTable('Resize Table', sheet, existing, replacement),

  /**
   * Grow a table for an entry typed just past its edge, bolding the header cells the growth adds
   * (header row when growing right, header column when growing down). Null when not adjacent.
   */
  expandTableForEdit(
    workbook: Workbook,
    sheet: string,
    table: Table,
    editedRow: number,
    editedColumn: number,
  ): EditCommand | null {
    const grownRange = TableAutoExpand.expandFor(table, editedRow, editedColumn);
    if (!grownRange) return null;
    const r = table.range;
    const steps: EditCommand[] = [
      DataEditCommands.replaceTable(sheet, table, { ...table, range: grownRange }),
    ];
    const headerStyle = (range: CellRange) =>
      EditCommands.applyStyle(workbook, sheet, range, styleTableHeader, 'Expand Table');
    if (table.hasHeader && grownRange.right > r.right) {
      steps.push(
        headerStyle(CellRange.fromBounds(grownRange.top, r.right + 1, grownRange.top, grownRange.right)),
      );
    }
    if (table.hasHeaderColumn && grownRange.bottom > r.bottom) {
      steps.push(
        headerStyle(CellRange.fromBounds(r.bottom + 1, grownRange.left, grownRange.bottom, grownRange.left)),
      );
    }
    return EditCommands.composite('Expand Table', steps);
  },

  // ---- Validation -------------------------------------------------------

  setValidation: (sheet: string, validation: DataValidation): EditCommand =>
    new ActionCommand(
      'Data Validation',
      (wb) => {
        requireSheet(wb, sheet).addValidation(validation);
        return EditResult.none;
      },
      (wb) => {
        requireSheet(wb, sheet).removeValidation(validation);
        return EditResult.none;
      },
    ),
} as const;

function styleTableHeader(style: CellStyle): CellStyle {
  return { ...style, font: { ...style.font, bold: true } };
}

function swapTable(label: string, sheet: string, existing: Table, replacement: Table): EditCommand {
  return new ActionCommand(
    label,
    (wb) => {
      requireSheet(wb, sheet).replaceTable(existing, replacement);
      return EditResult.none;
    },
    (wb) => {
      requireSheet(wb, sheet).replaceTable(replacement, existing);
      return EditResult.none;
    },
  );
}

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

/**
 * Changes nothing; keeps which rows of every filter that overlaps `range` (the sheet's filter and
 * table filters) are hidden, and puts exactly those flags back on undo.
 */
/**
 * The command `build` makes for `locale`, or, without one, for the workbook's document locale (its
 * effective locale) as it is when the command first applies.
 */
function inDocumentLocale(
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

function filteredRowsCommand(sheet: string, range: CellRange): EditCommand {
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
function rowVisibilityCommand(
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

function restoreLines(ws: Worksheet, rows: readonly LineState[], columns: readonly LineState[]): void {
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
