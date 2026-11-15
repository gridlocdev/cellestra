import { TableAutoExpand } from '../data/tables.ts';
import type { Table } from '../data/types.ts';
import { CellRange } from '../model/cellRange.ts';
import type { Workbook } from '../model/workbook.ts';
import type { CellStyle } from '../styling/cellStyle.ts';
import { CellValue } from '../values/cellValue.ts';
import { CellSnapshotCommand } from './cellSnapshotCommand.ts';
import { ActionCommand, EditResult, requireSheet, type EditCommand } from './editCommand.ts';
import { EditCommands } from './editCommands.ts';

/** Undoable table edits: create, restyle, resize, grow for an entry, remove. */
export const TableCommands = {
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
