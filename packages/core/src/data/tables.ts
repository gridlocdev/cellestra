import { CellRange } from '../model/cellRange.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import type { Table } from './types.ts';

/** A table's column labels: header text, else `Column{n}` (1-based). */
export function tableHeaderLabels(table: Table, sheet: Worksheet): string[] {
  const labels: string[] = [];
  for (let c = 0; c < table.range.columnCount; c++) {
    const v = table.hasHeader ? sheet.getValue(table.range.top, table.range.left + c) : null;
    labels.push(v?.kind === 'text' ? v.value : `Column${c + 1}`);
  }
  return labels;
}

/** The table called `name` (ignoring case) and its sheet, or null. Table names are unique in a workbook. */
export function findTableByName(workbook: Workbook, name: string): { sheet: Worksheet; table: Table } | null {
  const lower = name.trim().toLowerCase();
  for (const sheet of workbook.worksheets) {
    const table = sheet.tables.find((t) => t.name.toLowerCase() === lower);
    if (table) return { sheet, table };
  }
  return null;
}

/** The same table over a new range. */
export function tableWithRange(table: Table, range: CellRange): Table {
  return { ...table, range };
}

/** Tables grow when a value is typed just past their edge (PRD §8.10). */
export const TableAutoExpand = {
  /** The grown range for an entry at (row, column): directly below within the columns, or directly right within the rows. */
  expandFor(table: Table, row: number, column: number): CellRange | null {
    const r = table.range;
    const inColumns = column >= r.left && column <= r.right;
    const inRows = row >= r.top && row <= r.bottom;
    if (row === r.bottom + 1 && inColumns) return CellRange.fromBounds(r.top, r.left, row, r.right);
    if (column === r.right + 1 && inRows) return CellRange.fromBounds(r.top, r.left, r.bottom, column);
    return null;
  },
} as const;
