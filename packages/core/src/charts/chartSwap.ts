import type { ChartDefinition } from '@cellestra/charts';
import { ChartTypes, cloneChart, createSeries } from '@cellestra/charts';
import { quoteSheetIfNeeded } from '../model/cellAddress.ts';
import { CellRange } from '../model/cellRange.ts';
import type { Workbook } from '../model/workbook.ts';
import { fallbackSheetName } from './chartTableBinding.ts';
import { parseChartRange } from './chartReference.ts';

/** A chart with its rows and columns swapped, or why it can't be. */
export type ChartSwapResult =
  { readonly ok: true; readonly chart: ChartDefinition } | { readonly ok: false; readonly reason: string };

/**
 * Series laid out as columns (each series one column, the categories down the rows) or as rows.
 * A series sits in one *lane* (its column, or its row) and runs *along* the other direction.
 */
type Layout = 'columns' | 'rows';

/** The most series a swap makes: one per category, and a legend of more is unreadable. */
const MAX_SWAPPED_SERIES = 255;

const fail = (reason: string): ChartSwapResult => ({ ok: false, reason });

/**
 * Swap Rows and Columns: the series become the categories and the categories become the series,
 * as when Excel switches row and column. Works on a category chart whose series are side by side,
 * each one column (or one row) of the same cells; the category labels, when set, name the new
 * series, and the series' name cells, when they sit in one line beside the data, label the new
 * categories. Swapping twice gives the chart back.
 */
export const ChartSwap = {
  swap(workbook: Workbook, chart: ChartDefinition): ChartSwapResult {
    if (chart.source)
      return fail(
        'A chart plotting a table shows each column as a series. Switch to Range to swap rows and columns.',
      );
    if (ChartTypes.isXY(chart.type))
      return fail('Scatter and bubble charts plot X against Y, so they have no rows and columns to swap.');
    const values = chart.series.map((s) => parseChartRange(s.valuesReference));
    const first = values[0];
    if (!first || values.some((r) => r === null))
      return fail('Every series needs a range of cells before rows and columns can swap.');
    const ranges = values as CellRange[];
    const fallback = fallbackSheetName(workbook);
    const sheet = first.sheet ?? fallback;
    const onSheet = (r: CellRange) => (r.sheet ?? fallback).toLowerCase() === sheet.toLowerCase();
    if (!ranges.every(onSheet))
      return fail('Every series must be on the same sheet to swap rows and columns.');

    const fits = (layout: Layout): boolean =>
      ranges.every((r, i) =>
        layout === 'columns'
          ? r.left === r.right &&
            r.left === first.left + i &&
            r.top === first.top &&
            r.bottom === first.bottom
          : r.top === r.bottom && r.top === first.top + i && r.left === first.left && r.right === first.right,
      );
    const layout: Layout | null = fits('columns') ? 'columns' : fits('rows') ? 'rows' : null;
    if (!layout)
      return fail('To swap, each series must be one column (or one row) of the same length, side by side.');

    const lane = (r: CellRange) => (layout === 'columns' ? r.left : r.top);
    const alongFrom = layout === 'columns' ? first.top : first.left;
    const alongTo = layout === 'columns' ? first.bottom : first.right;
    const laneFrom = lane(first);
    const laneTo = laneFrom + ranges.length - 1;
    /** A reference to lanes `l0..l1` at positions `a0..a1` along them. */
    const ref = (l0: number, l1: number, a0: number, a1: number): string => {
      const r =
        layout === 'columns' ? CellRange.fromBounds(a0, l0, a1, l1) : CellRange.fromBounds(l0, a0, l1, a1);
      return `${quoteSheetIfNeeded(sheet)}.${r.toString()}`;
    };

    // The categories run along the series in a lane of their own.
    let categoryLane: number | null = null;
    if (chart.categoriesReference?.trim()) {
      const c = parseChartRange(chart.categoriesReference);
      const lines =
        c &&
        onSheet(c) &&
        (layout === 'columns'
          ? c.left === c.right && c.top === alongFrom && c.bottom === alongTo
          : c.top === c.bottom && c.left === alongFrom && c.right === alongTo);
      if (!c || !lines)
        return fail('The category labels must line up with the series to swap rows and columns.');
      categoryLane = lane(c);
    }

    // The series' name cells become the new categories when they sit in one line across the lanes.
    const names = chart.series.map((s) => parseChartRange(s.nameReference));
    const at = (r: CellRange) => (layout === 'columns' ? r.top : r.left);
    const nameAt = names[0] ? at(names[0]) : null;
    const namesInLine =
      nameAt !== null &&
      names.every(
        (n, i) => n !== null && n.isSingleCell && onSheet(n) && at(n) === nameAt && lane(n) === laneFrom + i,
      );

    const count = alongTo - alongFrom + 1;
    if (count > MAX_SWAPPED_SERIES)
      return fail(
        `Swapping would make ${count.toLocaleString('en-US')} series. Swap works up to ${MAX_SWAPPED_SERIES}.`,
      );

    const next = cloneChart(chart);
    next.categoriesReference = namesInLine ? ref(laneFrom, laneTo, nameAt, nameAt) : null;
    next.series = [];
    for (let a = alongFrom; a <= alongTo; a++)
      next.series.push(
        createSeries({
          valuesReference: ref(laneFrom, laneTo, a, a),
          nameReference: categoryLane === null ? null : ref(categoryLane, categoryLane, a, a),
        }),
      );
    return { ok: true, chart: next };
  },
} as const;
