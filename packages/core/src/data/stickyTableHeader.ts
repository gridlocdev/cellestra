import type { Worksheet } from '../model/worksheet.ts';

/** Sticky table headers: the column header shows the table's title once its header row scrolls away. */
export const StickyTableHeader = {
  /** The header row to borrow a label from for `column` with `topRow` at the top, or null. */
  labelRow(sheet: Worksheet, column: number, topRow: number): number | null {
    const t = sheet.findTable(topRow, column);
    return t && t.hasHeader && t.range.top < topRow ? t.range.top : null;
  },
} as const;
