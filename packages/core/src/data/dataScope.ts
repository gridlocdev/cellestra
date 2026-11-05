import type { CellRange } from '../model/cellRange.ts';
import type { Worksheet } from '../model/worksheet.ts';
import type { FilterSpec, SortSpec, Table } from './types.ts';

/**
 * Where a range's sort/filter state lives: on the table exactly covering it, else on the sheet's
 * AutoFilter when that is the range.
 */
export const DataScope = {
  tableAt: (sheet: Worksheet, range: CellRange): Table | null =>
    sheet.tables.find((t) => t.range.equals(range)) ?? null,

  effectiveFilter(sheet: Worksheet, range: CellRange): FilterSpec | null {
    const t = DataScope.tableAt(sheet, range);
    if (t) return t.filter ?? null;
    return sheet.activeFilter?.equals(range) ? sheet.activeFilterSpec : null;
  },

  effectiveSort(sheet: Worksheet, range: CellRange): SortSpec | null {
    const t = DataScope.tableAt(sheet, range);
    if (t) return t.sort ?? null;
    return sheet.activeFilter?.equals(range) ? sheet.activeSort : null;
  },
} as const;
