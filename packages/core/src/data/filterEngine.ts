import { CriteriaMatcher } from '../formulas/functions/criteriaMatcher.ts';
import type { CellRange } from '../model/cellRange.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { CellValue } from '../values/cellValue.ts';
import type { ColumnFilter, FilterSpec } from './types.ts';

/** Whether a value passes one column's criteria: in the allowed list (if any) and the condition (if any). */
export function columnFilterMatches(filter: ColumnFilter, value: CellValue, locale: string): boolean {
  return columnMatcher(filter, locale)(value);
}

/** One column's criteria as a predicate, its condition parsed once. */
function columnMatcher(filter: ColumnFilter, locale: string): (value: CellValue) => boolean {
  const allowed = filter.allowedValues ?? null;
  const condition =
    filter.condition != null ? CriteriaMatcher.build(CellValue.text(filter.condition), locale) : null;
  return (value) =>
    (!allowed || allowed.has(CellValue.toInvariantString(value))) && (!condition || condition.matches(value));
}

type RowTest = (sheet: Worksheet, row: number) => boolean;

/** The spec as one row predicate, each column's matcher built once (`exceptColumn` left out). */
function compile(spec: FilterSpec, exceptColumn: number | null, locale: string): RowTest {
  const tests: [column: number, matches: (value: CellValue) => boolean][] = [];
  for (const [column, filter] of spec.columns)
    if (column !== exceptColumn) tests.push([column, columnMatcher(filter, locale)]);
  return (sheet, row) => tests.every(([column, matches]) => matches(sheet.getValue(row, column)));
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
  /**
   * Mark the data rows that fail as filtered and clear the mark on the ones that pass. A row hidden
   * by hand stays hidden either way: the filter owns only its own flag. Rows below the sheet's last
   * stored cell hold no data and are left as they are, so a whole-column filter doesn't walk a
   * million rows; a row model is only created where a row's flag actually changes.
   */
  apply(sheet: Worksheet, range: CellRange, spec: FilterSpec, locale: string): void {
    const test = compile(spec, null, locale);
    const last = dataBottom(sheet, range);
    for (let row = firstDataRow(range, spec.hasHeader); row <= last; row++) {
      const filtered = !test(sheet, row);
      if (sheet.isRowFiltered(row) !== filtered) sheet.getOrCreateRow(row).filtered = filtered;
    }
  },

  /** Show the rows the filter hid; rows hidden by hand stay hidden. Only stored rows can be filtered. */
  clear(sheet: Worksheet, range: CellRange): void {
    for (const [row, model] of sheet.rows)
      if (row >= range.top && row <= range.bottom) model.filtered = false;
  },

  /** How many data rows the spec would leave visible — a preview that hides nothing. */
  countVisible(sheet: Worksheet, range: CellRange, spec: FilterSpec, locale: string): number {
    const test = compile(spec, null, locale);
    let visible = 0;
    for (let row = firstDataRow(range, spec.hasHeader); row <= dataBottom(sheet, range); row++) {
      if (test(sheet, row)) visible++;
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
    const others = otherFilters ? compile(otherFilters, column, locale) : null;
    for (let row = firstDataRow(range, hasHeader); row <= dataBottom(sheet, range); row++) {
      if (others && !others(sheet, row)) continue;
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

/** The range's last row that can hold data: rows past the last stored cell are all empty. */
function dataBottom(sheet: Worksheet, range: CellRange): number {
  return Math.min(range.bottom, sheet.cellExtent()?.lastRow ?? -1);
}
