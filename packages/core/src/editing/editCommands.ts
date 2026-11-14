import { CellRange } from '../model/cellRange.ts';
import type { Hyperlink } from '../model/hyperlink.ts';
import { MergedRegion } from '../model/mergedRegion.ts';
import type { Cell } from '../model/cell.ts';
import { cellKey, SheetDimensions } from '../model/sheetDimensions.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { CellStyles, NumberFormats, type CellStyle } from '../styling/cellStyle.ts';
import { StylePool } from '../styling/stylePool.ts';
import { CellValue } from '../values/cellValue.ts';
import { CellSnapshotCommand } from './cellSnapshotCommand.ts';
import {
  ActionCommand,
  CompositeEditCommand,
  EditResult,
  requireSheet,
  type EditCommand,
  type Position,
} from './editCommand.ts';
import { InputParser } from './inputParser.ts';

/**
 * A range larger than this (one full column's worth) is never iterated cell by cell: a whole-sheet
 * selection is clamped to the used range instead.
 */
export const OVERSIZED_CELL_CAP = 1 << 20;

/** The part of `range` inside the sheet's used range, or null (empty sheet, or no overlap). */
export function clampToUsed(sheet: Worksheet, range: CellRange): CellRange | null {
  const used = sheet.getUsedRange();
  if (!used) return null;
  const top = Math.max(range.top, used.top);
  const bottom = Math.min(range.bottom, used.bottom);
  const left = Math.max(range.left, used.left);
  const right = Math.min(range.right, used.right);
  if (top > bottom || left > right) return null;
  return CellRange.fromBounds(top, left, bottom, right);
}

/** The range to style: verbatim when modest, else clamped to the used range, else its top-left cell. */
function styleRange(sheet: Worksheet, range: CellRange): CellRange {
  if (range.rowCount * range.columnCount < OVERSIZED_CELL_CAP) return range;
  return clampToUsed(sheet, range) ?? new CellRange(range.topLeft.relative());
}

export function coversWholeColumns(range: CellRange): boolean {
  return range.top === 0 && range.bottom >= SheetDimensions.maxRows - 1;
}

export function coversWholeRows(range: CellRange): boolean {
  return range.left === 0 && range.right >= SheetDimensions.maxColumns - 1;
}

/** Factories for the everyday undoable edits. Every command names its sheet. */
export const EditCommands = {
  composite: (label: string, commands: readonly EditCommand[]): EditCommand =>
    new CompositeEditCommand(label, commands),

  // ---- Cell content -----------------------------------------------------

  setValue: (sheet: string, row: number, column: number, value: CellValue): EditCommand =>
    new CellSnapshotCommand('Set Value', sheet, [[row, column]], (ws) => ws.setValue(row, column, value)),

  /** `formula` is the source without its leading `=`; the cached value starts empty. */
  setFormula: (sheet: string, row: number, column: number, formula: string): EditCommand =>
    new CellSnapshotCommand('Set Formula', sheet, [[row, column]], (ws) =>
      ws.setFormula(row, column, formula, CellValue.empty),
    ),

  /**
   * Typed input, parsed in `locale`. With a workbook, a typed date also stamps its date format onto
   * a cell whose format is still General.
   */
  setInput(
    sheet: string,
    row: number,
    column: number,
    text: string,
    locale: string,
    workbook: Workbook | null = null,
  ): EditCommand {
    const parsed = InputParser.parse(text, locale);
    return new CellSnapshotCommand('Edit Cell', sheet, [[row, column]], (ws) => {
      if (parsed.formula !== null) {
        ws.setFormula(row, column, parsed.formula, CellValue.empty);
        return;
      }
      ws.setValue(row, column, parsed.value);
      const format = parsed.suggestedFormat;
      if (workbook && format) {
        const cell = ws.getOrCreate(row, column);
        const style = workbook.styles.get(cell.styleId);
        if (NumberFormats.isGeneral(style.numberFormat)) {
          cell.styleId = workbook.styles.intern({ ...style, numberFormat: format });
        }
        ws.pruneIfBlank(row, column, cell);
      }
    });
  },

  setHyperlink: (
    _workbook: Workbook,
    sheet: string,
    row: number,
    column: number,
    hyperlink: Hyperlink | null,
  ): EditCommand =>
    new CellSnapshotCommand('Set Hyperlink', sheet, [[row, column]], (ws) => {
      const cell = ws.getOrCreate(row, column);
      cell.hyperlink = hyperlink;
      ws.pruneIfBlank(row, column, cell);
    }),

  /** Drop values and formulas; style and hyperlink stay. */
  clearContents: (sheet: string, range: CellRange): EditCommand =>
    CellSnapshotCommand.overRange(
      'Clear Contents',
      sheet,
      (ws) => clampToUsed(ws, range),
      (ws) => {
        const clamped = clampToUsed(ws, range);
        if (!clamped) return;
        for (const [row, column] of ws.cellsIn(clamped)) ws.setValue(row, column, CellValue.empty);
      },
    ),

  clearFormatting: (workbook: Workbook, sheet: string, range: CellRange): EditCommand =>
    EditCommands.applyStyle(workbook, sheet, range, () => CellStyles.default, 'Clear Formatting'),

  /** Content, style, hyperlink and the comment or shapes a file gave it — the cell is gone afterwards. */
  clearAll: (sheet: string, range: CellRange): EditCommand =>
    CellSnapshotCommand.overRange(
      'Clear All',
      sheet,
      (ws) => clampToUsed(ws, range),
      (ws) => {
        const clamped = clampToUsed(ws, range);
        if (!clamped) return;
        for (const [row, column, cell] of ws.cellsIn(clamped)) {
          cell.formula = null;
          cell.value = CellValue.empty;
          cell.styleId = StylePool.defaultId;
          cell.hyperlink = null;
          cell.unmodeled = null;
          ws.pruneIfBlank(row, column, cell);
        }
      },
    ),

  /** Write a rectangular block of literals with its top-left at (top, left). */
  setBlock(
    sheet: string,
    top: number,
    left: number,
    values: readonly (readonly CellValue[])[],
    label = 'Paste',
  ): EditCommand {
    const positions: Position[] = [];
    values.forEach((row, r) => row.forEach((_, c) => positions.push([top + r, left + c])));
    return new CellSnapshotCommand(label, sheet, positions, (ws) => {
      values.forEach((row, r) => row.forEach((v, c) => ws.setValue(top + r, left + c, v)));
    });
  },

  /**
   * Tick or clear checkboxes: write TRUE or FALSE at each position. A formula cell computes its
   * own value, so it is left alone.
   */
  setChecked(sheet: string, positions: readonly Position[], checked: boolean): EditCommand {
    return new CellSnapshotCommand(checked ? 'Check' : 'Uncheck', sheet, positions, (ws) => {
      for (const [row, column] of positions) {
        if (ws.find(row, column)?.formula) continue;
        ws.setValue(row, column, CellValue.boolean(checked));
      }
    });
  },

  /**
   * Turn a range into checkboxes: the checkbox style on every cell, and FALSE in the empty ones so
   * each shows an unticked box (as in Google Sheets). Other values keep their value and show as
   * usual. A whole-column or whole-row range fills only up to the sheet's used area.
   */
  insertCheckboxes(workbook: Workbook, sheet: string, range: CellRange): EditCommand {
    const fillRange = (ws: Worksheet): CellRange | null =>
      coversWholeColumns(range) || coversWholeRows(range) || range.cellCount > OVERSIZED_CELL_CAP
        ? clampToUsed(ws, range)
        : range;
    const style = EditCommands.applyStyle(
      workbook,
      sheet,
      range,
      (s) => ({ ...s, checkbox: true }),
      'Checkbox',
    );
    const fill = CellSnapshotCommand.overRange('Checkbox', sheet, fillRange, (ws) => {
      for (const [row, column] of fillRange(ws)?.positions() ?? []) {
        const cell = ws.find(row, column);
        if (!cell || (cell.formula === null && cell.value.kind === 'empty'))
          ws.setValue(row, column, CellValue.boolean(false));
      }
    });
    return EditCommands.composite('Insert Checkbox', [style, fill]);
  },

  /** Take the checkbox style off a range; the TRUE/FALSE values stay. */
  removeCheckboxes(workbook: Workbook, sheet: string, range: CellRange): EditCommand {
    return EditCommands.applyStyle(
      workbook,
      sheet,
      range,
      ({ checkbox: _, ...rest }) => rest,
      'Remove Checkbox',
    );
  },

  // ---- Styling ----------------------------------------------------------

  /**
   * Transform the style of every cell in the range. A whole-column or whole-row range styles the
   * line models (plus the cells that need a style of their own) instead of materializing a cell per
   * position.
   */
  applyStyle(
    workbook: Workbook,
    sheet: string,
    range: CellRange,
    change: (style: CellStyle) => CellStyle,
    label = 'Format Cells',
  ): EditCommand {
    if (coversWholeColumns(range) || coversWholeRows(range))
      return lineStyleCommand(sheet, range, change, label);
    return new CellSnapshotCommand(
      label,
      sheet,
      (ws) => [...styleRange(ws, range).positions()],
      (ws) => {
        for (const [row, column] of styleRange(ws, range).positions()) {
          const cell = ws.getOrCreate(row, column);
          cell.styleId = workbook.styles.intern(change(workbook.styles.get(cell.styleId)));
          ws.pruneIfBlank(row, column, cell);
        }
      },
    );
  },

  // ---- Layout -----------------------------------------------------------

  setColumnWidth(sheet: string, column: number, width: number): EditCommand {
    let previous = 0;
    return new ActionCommand(
      'Resize Column',
      (wb) => {
        const col = requireSheet(wb, sheet).getOrCreateColumn(column);
        previous = col.width;
        col.width = width;
        return EditResult.columns(sheet, column, column);
      },
      (wb) => {
        requireSheet(wb, sheet).getOrCreateColumn(column).width = previous;
        return EditResult.columns(sheet, column, column);
      },
    );
  },

  setRowHeight(sheet: string, row: number, height: number): EditCommand {
    let previous = 0;
    return new ActionCommand(
      'Resize Row',
      (wb) => {
        const model = requireSheet(wb, sheet).getOrCreateRow(row);
        previous = model.height;
        model.height = height;
        return EditResult.rows(sheet, row, row);
      },
      (wb) => {
        requireSheet(wb, sheet).getOrCreateRow(row).height = previous;
        return EditResult.rows(sheet, row, row);
      },
    );
  },

  setFrozenPanes(sheet: string, rows: number, columns: number): EditCommand {
    let previousRows = 0;
    let previousColumns = 0;
    return new ActionCommand(
      'Freeze Panes',
      (wb) => {
        const ws = requireSheet(wb, sheet);
        previousRows = ws.frozenRows;
        previousColumns = ws.frozenColumns;
        ws.frozenRows = rows;
        ws.frozenColumns = columns;
        return EditResult.none;
      },
      (wb) => {
        const ws = requireSheet(wb, sheet);
        ws.frozenRows = previousRows;
        ws.frozenColumns = previousColumns;
        return EditResult.none;
      },
    );
  },

  setColumnsHidden(sheet: string, firstColumn: number, lastColumn: number, hidden: boolean): EditCommand {
    const previous: boolean[] = [];
    return new ActionCommand(
      hidden ? 'Hide Columns' : 'Unhide Columns',
      (wb) => {
        const ws = requireSheet(wb, sheet);
        for (let c = firstColumn; c <= lastColumn; c++) {
          const model = ws.getOrCreateColumn(c);
          previous[c - firstColumn] = model.hidden;
          model.hidden = hidden;
        }
        return EditResult.columns(sheet, firstColumn, lastColumn);
      },
      (wb) => {
        const ws = requireSheet(wb, sheet);
        for (let c = firstColumn; c <= lastColumn; c++)
          ws.getOrCreateColumn(c).hidden = previous[c - firstColumn]!;
        return EditResult.columns(sheet, firstColumn, lastColumn);
      },
    );
  },

  setRowsHidden(sheet: string, firstRow: number, lastRow: number, hidden: boolean): EditCommand {
    const previous: boolean[] = [];
    return new ActionCommand(
      hidden ? 'Hide Rows' : 'Unhide Rows',
      (wb) => {
        const ws = requireSheet(wb, sheet);
        for (let r = firstRow; r <= lastRow; r++) {
          const model = ws.getOrCreateRow(r);
          previous[r - firstRow] = model.manuallyHidden;
          model.manuallyHidden = hidden;
        }
        return EditResult.rows(sheet, firstRow, lastRow);
      },
      (wb) => {
        const ws = requireSheet(wb, sheet);
        for (let r = firstRow; r <= lastRow; r++)
          ws.getOrCreateRow(r).manuallyHidden = previous[r - firstRow]!;
        return EditResult.rows(sheet, firstRow, lastRow);
      },
    );
  },

  // ---- Merge ------------------------------------------------------------

  mergeCells: (sheet: string, range: CellRange): EditCommand =>
    new ActionCommand(
      'Merge Cells',
      (wb) => {
        requireSheet(wb, sheet).addMerge(new MergedRegion(range));
        return EditResult.rows(sheet, range.top, range.bottom);
      },
      (wb) => {
        requireSheet(wb, sheet).removeMerge(range.top, range.left);
        return EditResult.rows(sheet, range.top, range.bottom);
      },
    ),

  unmergeCells(sheet: string, row: number, column: number): EditCommand {
    let removed: MergedRegion | null = null;
    return new ActionCommand(
      'Unmerge Cells',
      (wb) => {
        const ws = requireSheet(wb, sheet);
        removed = ws.findMerge(row, column) ?? null;
        if (removed) ws.removeMerge(row, column);
        return mergeRows(sheet, removed);
      },
      (wb) => {
        if (removed) requireSheet(wb, sheet).addMerge(removed);
        return mergeRows(sheet, removed);
      },
    );
  },
} as const;

type StyleChange = (style: CellStyle) => CellStyle;

/** The rows a merge covers, or nothing when there was none. */
function mergeRows(sheet: string, merge: MergedRegion | null): EditResult {
  return merge ? EditResult.rows(sheet, merge.range.top, merge.range.bottom) : EditResult.none;
}

/**
 * Style whole columns and/or whole rows through their line models, never one cell per position.
 * Every position in the range ends up looking like `change` of how it looked before: lines take the
 * transform, cells with a style of their own take it directly, and where the crossing line's style
 * used to win (a styled row under a column style: cell → row → column) the crossing gets a cell of
 * its own, so the new style shows there too (as LibreOffice and Excel do).
 */
export function lineStyleCommand(
  sheet: string,
  range: CellRange,
  change: StyleChange,
  label: string,
): EditCommand {
  const columns = coversWholeColumns(range);
  const rows = coversWholeRows(range);
  let rowsBefore: [index: number, styleId: number][] = [];
  let columnsBefore: [index: number, styleId: number][] = [];
  let cellsBefore: [row: number, column: number, before: Cell | null][] = [];
  // Whole columns (the whole sheet included) or whole rows: every cell in them may look different.
  const styledLines = columns
    ? EditResult.columns(sheet, range.left, range.right)
    : EditResult.rows(sheet, range.top, range.bottom);
  return new ActionCommand(
    label,
    (wb) => {
      const ws = requireSheet(wb, sheet);
      const memo = new Map<number, number>();
      const restyled = (id: number): number => {
        let out = memo.get(id);
        if (out === undefined) {
          out = wb.styles.intern(change(wb.styles.get(id)));
          memo.set(id, out);
        }
        return out;
      };
      rowsBefore = [];
      columnsBefore = [];
      cellsBefore = [];

      // Work out the cells first: they depend on the line styles as they were.
      const cellStyles = new Map<number, [row: number, column: number, styleId: number]>();
      const setCell = (row: number, column: number, styleId: number) =>
        cellStyles.set(cellKey(row, column), [row, column, styleId]);
      for (const [row, column, cell] of ws.enumerateCells()) {
        if (range.contains(row, column) && cell.styleId !== StylePool.defaultId)
          setCell(row, column, restyled(cell.styleId));
      }
      const own = (row: number, column: number): number =>
        ws.find(row, column)?.styleId ?? StylePool.defaultId;
      if (columns && !rows) {
        // A styled row beats the column: its crossings need the change as cells of their own.
        for (const [row, model] of ws.rows) {
          if (model.styleId === StylePool.defaultId) continue;
          for (let column = range.left; column <= range.right; column++)
            if (own(row, column) === StylePool.defaultId) setCell(row, column, restyled(model.styleId));
        }
      } else if (rows && !columns) {
        // A row with no style of its own showed its columns' styles; the new row style would hide them.
        for (const [column, model] of ws.columns) {
          if (model.styleId === StylePool.defaultId || column < range.left || column > range.right) continue;
          for (let row = range.top; row <= range.bottom; row++) {
            if (ws.findRow(row)?.styleId ?? StylePool.defaultId) continue;
            if (own(row, column) === StylePool.defaultId) setCell(row, column, restyled(model.styleId));
          }
        }
      }

      if (columns) {
        for (let column = range.left; column <= range.right; column++) {
          const model = ws.getOrCreateColumn(column);
          columnsBefore.push([column, model.styleId]);
          model.styleId = restyled(model.styleId);
        }
      }
      if (rows && !columns) {
        for (let row = range.top; row <= range.bottom; row++) {
          const model = ws.getOrCreateRow(row);
          rowsBefore.push([row, model.styleId]);
          model.styleId = restyled(model.styleId);
        }
      } else if (rows) {
        // The whole sheet: the columns cover unstyled rows; styled rows take the change themselves.
        for (const [row, model] of ws.rows) {
          if (model.styleId === StylePool.defaultId) continue;
          rowsBefore.push([row, model.styleId]);
          model.styleId = restyled(model.styleId);
        }
      }
      for (const [row, column, styleId] of cellStyles.values()) {
        cellsBefore.push([row, column, ws.find(row, column)?.clone() ?? null]);
        const cell = ws.getOrCreate(row, column);
        cell.styleId = styleId;
        ws.pruneIfBlank(row, column, cell);
      }
      return styledLines; // a style change moves no values: nothing to recalc
    },
    (wb) => {
      const ws = requireSheet(wb, sheet);
      for (const [column, styleId] of columnsBefore) ws.getOrCreateColumn(column).styleId = styleId;
      for (const [row, styleId] of rowsBefore) ws.getOrCreateRow(row).styleId = styleId;
      for (const [row, column, before] of cellsBefore) ws.putCell(row, column, before?.clone() ?? null);
      return styledLines;
    },
  );
}
