import type { Worksheet } from '../../src/model/worksheet.ts';
import type { CellValue } from '../../src/values/cellValue.ts';

/** The text of a text value, else null. */
export function asText(v: CellValue): string | null {
  return v.kind === 'text' ? v.value : null;
}

/** The number of a number value, else null. */
export function asNumber(v: CellValue): number | null {
  return v.kind === 'number' ? v.value : null;
}

export const textAt = (sheet: Worksheet, row: number, column: number): string | null =>
  asText(sheet.getValue(row, column));
export const numberAt = (sheet: Worksheet, row: number, column: number): number | null =>
  asNumber(sheet.getValue(row, column));
