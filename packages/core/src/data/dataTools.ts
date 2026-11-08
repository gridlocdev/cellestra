import { rewriteFormula } from '../editing/formulaRewrite.ts';
import { ReferenceAdjuster } from '../formulas/references/referenceAdjuster.ts';
import type { Cell } from '../model/cell.ts';
import type { CellRange } from '../model/cellRange.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { InputParser } from '../editing/inputParser.ts';
import { getHostLocale } from '../model/documentSettings.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
import { parseCsv } from '../serialization/csv/csvReader.ts';
import { NumberFormats } from '../styling/cellStyle.ts';
import type { StylePool } from '../styling/stylePool.ts';
import { CellValue } from '../values/cellValue.ts';

/** Remove Duplicates (PRD §8.11). */
export const DuplicateRemover = {
  /**
   * Drop data rows whose key columns (all columns when `keyColumns` is empty) repeat an earlier row,
   * compacting survivors upward within the range. A moved row's relative references into the range
   * follow it, as they do when sorting. Returns how many rows went.
   */
  remove(sheet: Worksheet, range: CellRange, keyColumns: readonly number[], hasHeader: boolean): number {
    const first = range.top + (hasHeader ? 1 : 0);
    if (first > range.bottom) return 0;
    const keys =
      keyColumns.length > 0
        ? keyColumns
        : Array.from({ length: range.columnCount }, (_, i) => range.left + i);
    const seen = new Set<string>();
    const survivors: { originalRow: number; cells: (Cell | null)[] }[] = [];
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
      survivors.push({ originalRow: row, cells });
    }
    let writeRow = first;
    for (const { originalRow, cells } of survivors) {
      const rowDelta = writeRow - originalRow;
      cells.forEach((cell, c) => {
        if (cell && cell.formula !== null && rowDelta !== 0) {
          cell.formula = rewriteFormula(cell.formula, (expr) =>
            ReferenceAdjuster.forSortMove(expr, rowDelta, first, range.bottom),
          );
        }
        sheet.putCell(writeRow, range.left + c, cell);
      });
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
   * right; an empty field clears its cell. Each field is read as typed input in `locale` (a number,
   * a date, TRUE/FALSE, else text; never a formula), and with `styles` a date also gets its date
   * format where the cell's format is General. Returns the widest field count (at least 1).
   */
  split(
    sheet: Worksheet,
    sourceColumn: CellRange,
    delimiter: string,
    options: { readonly locale?: string; readonly styles?: StylePool | null } = {},
  ): number {
    const column = sourceColumn.left;
    const locale = options.locale ?? getHostLocale();
    let maxColumns = 1;
    for (let row = sourceColumn.top; row <= sourceColumn.bottom; row++) {
      const value = sheet.getValue(row, column);
      if (value.kind !== 'text') continue;
      const fields = splitFields(value.value, delimiter);
      maxColumns = Math.max(maxColumns, fields.length);
      fields.forEach((f, i) => {
        if (column + i >= SheetDimensions.maxColumns) return;
        writeField(sheet, row, column + i, f, locale, options.styles ?? null);
      });
    }
    return maxColumns;
  },

  /** The fields one cell's text splits into (the first CSV record). */
  fields: (text: string, delimiter: string): string[] => splitFields(text, delimiter),
} as const;

function writeField(
  sheet: Worksheet,
  row: number,
  column: number,
  field: string,
  locale: string,
  styles: StylePool | null,
): void {
  if (field.length === 0) {
    sheet.setValue(row, column, CellValue.empty);
    return;
  }
  // A field that would read as a formula, or carries the text marker, stays the text it was.
  const parsed = field[0] === '=' || field[0] === "'" ? null : InputParser.parse(field, locale);
  if (!parsed || parsed.formula !== null) {
    sheet.setValue(row, column, CellValue.text(field));
    return;
  }
  sheet.setValue(row, column, parsed.value);
  const format = parsed.suggestedFormat;
  if (styles && format) {
    const cell = sheet.getOrCreate(row, column);
    const style = styles.get(cell.styleId);
    if (NumberFormats.isGeneral(style.numberFormat))
      cell.styleId = styles.intern({ ...style, numberFormat: format });
  }
}

function splitFields(text: string, delimiter: string): string[] {
  return parseCsv(text, { delimiter })[0] ?? [];
}
