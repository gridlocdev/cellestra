import { CellRange } from '../model/cellRange.ts';
import type { Hyperlink } from '../model/hyperlink.ts';
import { MergedRegion } from '../model/mergedRegion.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
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

export function clampedPositions(sheet: Worksheet, range: CellRange): Position[] {
  const clamped = clampToUsed(sheet, range);
  return clamped ? [...clamped.positions()] : [];
}

/** The range to style: verbatim when modest, else clamped to the used range, else its top-left cell. */
function styleRange(sheet: Worksheet, range: CellRange): CellRange {
  if (range.rowCount * range.columnCount < OVERSIZED_CELL_CAP) return range;
  return clampToUsed(sheet, range) ?? new CellRange(range.topLeft.relative());
}

function coversWholeColumns(range: CellRange): boolean {
  return range.top === 0 && range.bottom >= SheetDimensions.maxRows - 1;
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
    new CellSnapshotCommand(
      'Clear Contents',
      sheet,
      (ws) => clampedPositions(ws, range),
      (ws) => {
        for (const [row, column] of clampToUsed(ws, range)?.positions() ?? []) {
          if (ws.find(row, column)) ws.setValue(row, column, CellValue.empty);
        }
      },
    ),

  clearFormatting: (workbook: Workbook, sheet: string, range: CellRange): EditCommand =>
    EditCommands.applyStyle(workbook, sheet, range, () => CellStyles.default, 'Clear Formatting'),

  /** Content, style, and hyperlink — the cell is gone afterwards. */
  clearAll: (sheet: string, range: CellRange): EditCommand =>
    new CellSnapshotCommand(
      'Clear All',
      sheet,
      (ws) => clampedPositions(ws, range),
      (ws) => {
        for (const [row, column] of clampToUsed(ws, range)?.positions() ?? []) {
          const cell = ws.find(row, column);
          if (!cell) continue;
          cell.formula = null;
          cell.value = CellValue.empty;
          cell.styleId = StylePool.defaultId;
          cell.hyperlink = null;
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

  // ---- Styling ----------------------------------------------------------

  /**
   * Transform the style of every cell in the range. A whole-column range styles the column models
   * (plus cells carrying a style of their own) instead of materializing a million cells.
   */
  applyStyle(
    workbook: Workbook,
    sheet: string,
    range: CellRange,
    change: (style: CellStyle) => CellStyle,
    label = 'Format Cells',
  ): EditCommand {
    if (coversWholeColumns(range)) return applyColumnStyle(workbook, sheet, range, change, label);
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
        return EditResult.none;
      },
      (wb) => {
        requireSheet(wb, sheet).getOrCreateColumn(column).width = previous;
        return EditResult.none;
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
        return EditResult.none;
      },
      (wb) => {
        requireSheet(wb, sheet).getOrCreateRow(row).height = previous;
        return EditResult.none;
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
        return EditResult.none;
      },
      (wb) => {
        const ws = requireSheet(wb, sheet);
        for (let c = firstColumn; c <= lastColumn; c++)
          ws.getOrCreateColumn(c).hidden = previous[c - firstColumn]!;
        return EditResult.none;
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
          previous[r - firstRow] = model.hidden;
          model.hidden = hidden;
        }
        return EditResult.none;
      },
      (wb) => {
        const ws = requireSheet(wb, sheet);
        for (let r = firstRow; r <= lastRow; r++) ws.getOrCreateRow(r).hidden = previous[r - firstRow]!;
        return EditResult.none;
      },
    );
  },

  // ---- Merge ------------------------------------------------------------

  mergeCells: (sheet: string, range: CellRange): EditCommand =>
    new ActionCommand(
      'Merge Cells',
      (wb) => {
        requireSheet(wb, sheet).addMerge(new MergedRegion(range));
        return EditResult.none;
      },
      (wb) => {
        requireSheet(wb, sheet).removeMerge(range.top, range.left);
        return EditResult.none;
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
        return EditResult.none;
      },
      (wb) => {
        if (removed) requireSheet(wb, sheet).addMerge(removed);
        return EditResult.none;
      },
    );
  },
} as const;

function applyColumnStyle(
  workbook: Workbook,
  sheet: string,
  range: CellRange,
  change: (style: CellStyle) => CellStyle,
  label: string,
): EditCommand {
  let columnsBefore: [column: number, styleId: number][] = [];
  let cellsBefore: [row: number, column: number, styleId: number][] = [];
  return new ActionCommand(
    label,
    (wb) => {
      const ws = requireSheet(wb, sheet);
      columnsBefore = [];
      cellsBefore = [];
      for (let column = range.left; column <= range.right; column++) {
        const model = ws.getOrCreateColumn(column);
        columnsBefore.push([column, model.styleId]);
        model.styleId = workbook.styles.intern(change(workbook.styles.get(model.styleId)));
      }
      // Cells with a style of their own aren't covered by the column default, so they take the
      // transform directly; unstyled ones inherit it and are left alone.
      for (const [row, column, cell] of [...ws.enumerateCells()]) {
        if (column < range.left || column > range.right || cell.styleId === StylePool.defaultId) continue;
        cellsBefore.push([row, column, cell.styleId]);
        cell.styleId = workbook.styles.intern(change(workbook.styles.get(cell.styleId)));
      }
      return EditResult.none; // a style change moves no values: nothing to recalc
    },
    (wb) => {
      const ws = requireSheet(wb, sheet);
      for (const [column, styleId] of columnsBefore) ws.getOrCreateColumn(column).styleId = styleId;
      for (const [row, column, styleId] of cellsBefore) {
        const cell = ws.find(row, column);
        if (cell) cell.styleId = styleId;
      }
      return EditResult.none;
    },
  );
}
