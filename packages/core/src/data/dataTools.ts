import type { Cell } from '../model/cell.ts';
import type { CellRange } from '../model/cellRange.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { parseCsv } from '../serialization/csv/csvReader.ts';
import { CellValue } from '../values/cellValue.ts';

/** Remove Duplicates (PRD §8.11). */
export const DuplicateRemover = {
  /**
   * Drop data rows whose key columns (all columns when `keyColumns` is empty) repeat an earlier row,
   * compacting survivors upward within the range. Returns how many rows went.
   */
  remove(sheet: Worksheet, range: CellRange, keyColumns: readonly number[], hasHeader: boolean): number {
    const first = range.top + (hasHeader ? 1 : 0);
    if (first > range.bottom) return 0;
    const keys =
      keyColumns.length > 0
        ? keyColumns
        : Array.from({ length: range.columnCount }, (_, i) => range.left + i);
    const seen = new Set<string>();
    const survivors: (Cell | null)[][] = [];
    let removed = 0;
    for (let row = first; row <= range.bottom; row++) {
      const key = keys.map((c) => CellValue.toInvariantString(sheet.getValue(row, c))).join('\u001f');
      if (seen.has(key)) {
        removed++;
        continue;
      }
      seen.add(key);
      const cells: (Cell | null)[] = [];
      for (let c = 0; c < range.columnCount; c++)
        cells.push(sheet.find(row, range.left + c)?.clone() ?? null);
      survivors.push(cells);
    }
    let writeRow = first;
    for (const cells of survivors) {
      cells.forEach((cell, c) => sheet.putCell(writeRow, range.left + c, cell));
      writeRow++;
    }
    for (; writeRow <= range.bottom; writeRow++) {
      for (let c = 0; c < range.columnCount; c++) sheet.remove(writeRow, range.left + c);
    }
    return removed;
  },
} as const;

/** Text to Columns (PRD §8.11). */
export const TextToColumns = {
  /**
   * Split each text cell of the source column on `delimiter` (CSV quoting rules) into the cells to its
   * right; an empty field clears its cell. Returns the widest field count (at least 1).
   */
  split(sheet: Worksheet, sourceColumn: CellRange, delimiter: string): number {
    const column = sourceColumn.left;
    let maxColumns = 1;
    for (let row = sourceColumn.top; row <= sourceColumn.bottom; row++) {
      const value = sheet.getValue(row, column);
      if (value.kind !== 'text') continue;
      const fields = splitFields(value.value, delimiter);
      maxColumns = Math.max(maxColumns, fields.length);
      fields.forEach((f, i) =>
        sheet.setValue(row, column + i, f.length === 0 ? CellValue.empty : CellValue.text(f)),
      );
    }
    return maxColumns;
  },

  /** The fields one cell's text splits into (the first CSV record). */
  fields: (text: string, delimiter: string): string[] => splitFields(text, delimiter),
} as const;

function splitFields(text: string, delimiter: string): string[] {
  return parseCsv(text, { delimiter })[0] ?? [];
}
