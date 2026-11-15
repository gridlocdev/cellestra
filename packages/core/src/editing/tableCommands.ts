import { DataRegion } from '../data/dataRegion.ts';
import { TableAutoExpand } from '../data/tables.ts';
import { table as makeTable, type Table } from '../data/types.ts';
import { CellRange } from '../model/cellRange.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import type { CellStyle } from '../styling/cellStyle.ts';
import { CellValue } from '../values/cellValue.ts';
import { CellSnapshotCommand } from './cellSnapshotCommand.ts';
import { ActionCommand, EditResult, requireSheet, type EditCommand } from './editCommand.ts';
import { EditCommands } from './editCommands.ts';
import { StructuralEditCommand } from './structuralEditCommand.ts';

/** How `TableCommands.create` treats the range's first row. */
export type TableHeaderChoice = boolean | 'labels';

/** A table ready to create, or why not: the range overlaps these tables. */
export type TablePlan =
  | { readonly ok: true; readonly command: EditCommand; readonly table: Table }
  | { readonly ok: false; readonly overlapping: readonly Table[] };

/** Undoable table edits: create, restyle, resize, grow for an entry, remove. */
export const TableCommands = {
  /**
   * Create Table, as one undo step. The name defaults to the next free `TableN` in the workbook
   * (tables are saved as document-wide database ranges). `header: 'labels'` gives every table a
   * header row: a first row of text labels is it, otherwise "Column N" labels go in the blank row
   * just above, or in a row inserted there. A range that overlaps a table is refused unless
   * `replace` is set, which removes those tables in the same step.
   */
  create(
    workbook: Workbook,
    sheet: string,
    range: CellRange,
    options: { name?: string | null; header?: TableHeaderChoice; replace?: boolean } = {},
  ): TablePlan {
    const ws = requireSheet(workbook, sheet);
    const overlapping = ws.tables.filter((t) => t.range.intersects(range));
    if (overlapping.length > 0 && options.replace !== true) return { ok: false, overlapping };
    const steps: EditCommand[] = overlapping.map((t) => TableCommands.removeTable(sheet, t));
    const header = options.header ?? true;
    let target = range;
    if (header === 'labels' && !DataRegion.firstRowIsLabels(ws, range)) {
      const above = range.top - 1;
      const free =
        above >= 0 &&
        rowIsBlank(ws, above, range.left, range.right) &&
        !ws.tables.some((t) => t.range.top <= above && t.range.bottom >= above);
      if (free) target = CellRange.fromBounds(above, range.left, range.bottom, range.right);
      else {
        steps.push(new StructuralEditCommand(sheet, 'row', 'insert', range.top, 1));
        target = CellRange.fromBounds(range.top, range.left, range.bottom + 1, range.right);
      }
    }
    const table = makeTable(options.name ?? TableCommands.nextName(workbook), target, {
      hasHeader: header !== false,
    });
    steps.push(TableCommands.createStyledTable(workbook, sheet, table));
    const command = steps.length === 1 ? steps[0]! : EditCommands.composite('Create Table', steps);
    return { ok: true, command, table };
  },

  /** The first `TableN` no table in the workbook is called (names compare without case). */
  nextName(workbook: Workbook): string {
    const taken = new Set(workbook.worksheets.flatMap((s) => s.tables).map((t) => t.name.toLowerCase()));
    for (let n = 1; ; n++) if (!taken.has(`table${n}`)) return `Table${n}`;
  },

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
    const add = TableCommands.createTable(sheet, table);
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
      enabled ? styleTableHeader : TableCommands.clearTableHeader,
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
    const steps: EditCommand[] = [TableCommands.replaceTable(sheet, table, { ...table, range: grownRange })];
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

function rowIsBlank(sheet: Worksheet, row: number, left: number, right: number): boolean {
  for (let c = left; c <= right; c++) if (sheet.getValue(row, c).kind !== 'empty') return false;
  return true;
}
