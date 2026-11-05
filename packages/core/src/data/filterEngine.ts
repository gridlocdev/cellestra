import { CriteriaMatcher } from '../formulas/functions/criteriaMatcher.ts';
import type { CellRange } from '../model/cellRange.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { CellValue } from '../values/cellValue.ts';
import type { ColumnFilter, FilterSpec } from './types.ts';

/** Whether a value passes one column's criteria: in the allowed list (if any) and the condition (if any). */
export function columnFilterMatches(filter: ColumnFilter, value: CellValue, locale: string): boolean {
  if (filter.allowedValues && !filter.allowedValues.has(CellValue.toInvariantString(value))) return false;
  if (
    filter.condition != null &&
    !CriteriaMatcher.build(CellValue.text(filter.condition), locale).matches(value)
  ) {
    return false;
  }
  return true;
}

/** An empty-to-start filter spec over the given per-column criteria. */
export function filterSpec(columns: ReadonlyMap<number, ColumnFilter>, hasHeader = true): FilterSpec {
  return { columns, hasHeader };
}

/**
 * AutoFilter (PRD §8.12): hides data rows that fail any column's criteria. The header row (when the
 * spec has one) is never hidden nor offered as a value.
 */
export const FilterEngine = {
  apply(sheet: Worksheet, range: CellRange, spec: FilterSpec, locale: string): void {
    for (let row = firstDataRow(range, spec.hasHeader); row <= range.bottom; row++) {
      sheet.getOrCreateRow(row).hidden = !passes(sheet, row, spec, null, locale);
    }
  },

  /** Unhide every row of the range (only rows that have a stored model can be hidden). */
  clear(sheet: Worksheet, range: CellRange): void {
    for (let row = range.top; row <= range.bottom; row++) {
      const model = sheet.findRow(row);
      if (model) model.hidden = false;
    }
  },

  /** How many data rows the spec would leave visible — a preview that hides nothing. */
  countVisible(sheet: Worksheet, range: CellRange, spec: FilterSpec, locale: string): number {
    let visible = 0;
    for (let row = firstDataRow(range, spec.hasHeader); row <= range.bottom; row++) {
      if (passes(sheet, row, spec, null, locale)) visible++;
    }
    return visible;
  },

  /**
   * The column's distinct values in row order (deduplicated by invariant text). With `otherFilters`,
   * only rows passing every *other* column's criteria count — a column's own filter never narrows
   * its own list, so the user can widen it again.
   */
  distinctValues(
    sheet: Worksheet,
    range: CellRange,
    column: number,
    hasHeader: boolean,
    otherFilters: FilterSpec | null = null,
    locale = 'en-US',
  ): CellValue[] {
    const seen = new Set<string>();
    const out: CellValue[] = [];
    for (let row = firstDataRow(range, hasHeader); row <= range.bottom; row++) {
      if (otherFilters && !passes(sheet, row, otherFilters, column, locale)) continue;
      const value = sheet.getValue(row, column);
      const key = CellValue.toInvariantString(value);
      if (!seen.has(key)) {
        seen.add(key);
        out.push(value);
      }
    }
    return out;
  },
} as const;

function firstDataRow(range: CellRange, hasHeader: boolean): number {
  return range.top + (hasHeader ? 1 : 0);
}

function passes(
  sheet: Worksheet,
  row: number,
  spec: FilterSpec,
  exceptColumn: number | null,
  locale: string,
): boolean {
  for (const [column, filter] of spec.columns) {
    if (column === exceptColumn) continue;
    if (!columnFilterMatches(filter, sheet.getValue(row, column), locale)) return false;
  }
  return true;
}
