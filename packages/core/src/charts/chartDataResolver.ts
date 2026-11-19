import { labelOf } from './chartLabels.ts';
import type {
  ChartAggregate,
  ChartDefinition,
  ChartSeries,
  ChartStacking,
  ChartType,
} from '@cellestra/charts';
import { ChartTypes } from '@cellestra/charts';
import { Defaults } from '../defaults.ts';
import { quoteSheetIfNeeded } from '../model/cellAddress.ts';
import { CellRange } from '../model/cellRange.ts';
import { ChartGrouping, type GroupingSeries } from './chartGrouping.ts';
import { ChartTableBinding } from './chartTableBinding.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';

/** One series resolved to concrete numbers ready to plot. `NaN` marks a blank/non-numeric cell (a gap). */
export interface ResolvedSeries {
  readonly name: string;
  /** The effective type: the series override, or the chart's type. */
  readonly type: ChartType;
  readonly values: readonly number[];
  /** X values for an X/Y series (scatter/bubble), or null. */
  readonly xValues: readonly number[] | null;
  readonly bubbleSizes: readonly number[] | null;
}

/** A whole chart resolved to plottable data plus the presentation flags the renderer needs. */
export interface ResolvedChart {
  readonly title: string;
  readonly type: ChartType;
  readonly stacking: ChartStacking;
  readonly showLegend: boolean;
  readonly xAxisTitle: string | null;
  readonly yAxisTitle: string | null;
  readonly categories: readonly string[];
  readonly series: readonly ResolvedSeries[];
  /** How repeated category labels were combined; omitted means sum. */
  readonly aggregate?: ChartAggregate;
  /**
   * How many source rows each category merged (for hover details), or null/omitted when no rows
   * merged. Index `c` matches `categories[c]`.
   */
  readonly rowCounts?: readonly number[] | null;
  /** True when there is nothing to draw (no series, or every series is empty). */
  readonly isEmpty: boolean;
}

/** Qualify a sheet-less reference with the fallback sheet so it resolves to a real sheet. */
function qualify(reference: string, fallbackSheet: string): string {
  const r = reference.trim();
  return r.includes('.') ? r : `${quoteSheetIfNeeded(fallbackSheet)}.${r}`;
}

/** The sheet an unqualified reference targets: the first data (non-chart) sheet, else the first sheet. */
function defaultSheetName(workbook: Workbook): string {
  const data = workbook.worksheets.find((s) => !s.isChartSheet);
  return data?.name ?? workbook.worksheets[0]?.name ?? Defaults.firstSheetName;
}

function resolveRange(
  workbook: Workbook,
  reference: string | null | undefined,
  fallbackSheet: string,
): { range: CellRange; sheet: Worksheet } | null {
  if (!reference || reference.trim() === '') return null;
  const range = CellRange.tryParse(qualify(reference, fallbackSheet));
  if (!range) return null;
  const sheet = workbook.findSheet(range.sheet ?? fallbackSheet);
  if (!sheet) return null;
  // Trim to the used range (a whole column holds a million mostly-empty rows), then cap the cell count
  // so a half-typed reference like A1:XFD1048576 can't allocate billions of points.
  const used = sheet.getUsedRange();
  if (!used) return null;
  const bottom = Math.min(range.bottom, used.bottom);
  const right = Math.min(range.right, used.right);
  if (bottom < range.top || right < range.left) return null;
  const columns = right - range.left + 1;
  const rows = Math.min(bottom - range.top + 1, Math.max(1, Math.floor(MAX_CHART_CELLS / columns)));
  const cols = Math.min(columns, MAX_CHART_CELLS);
  return {
    range: CellRange.fromBounds(range.top, range.left, range.top + rows - 1, range.left + cols - 1),
    sheet,
  };
}

/** The most points one chart reference reads. */
const MAX_CHART_CELLS = 100_000;

function readCells(workbook: Workbook, reference: string | null, fallbackSheet: string): GroupingSeries {
  const resolved = resolveRange(workbook, reference, fallbackSheet);
  const values: number[] = [];
  const filled: boolean[] = [];
  if (!resolved) return { values, filled };
  for (const [row, column] of resolved.range.positions()) {
    const v = resolved.sheet.getValue(row, column);
    values.push(v.kind === 'number' ? v.value : v.kind === 'boolean' ? (v.value ? 1 : 0) : NaN);
    filled.push(v.kind !== 'empty');
  }
  return { values, filled };
}

function readNumbers(workbook: Workbook, reference: string | null, fallbackSheet: string): number[] {
  return [...readCells(workbook, reference, fallbackSheet).values];
}

function readLabels(workbook: Workbook, reference: string | null, fallbackSheet: string): string[] {
  const resolved = resolveRange(workbook, reference, fallbackSheet);
  if (!resolved) return [];
  return [...resolved.range.positions()].map(([row, column]) =>
    labelOf(resolved.sheet.getValue(row, column)),
  );
}

function seriesName(workbook: Workbook, series: ChartSeries, fallbackSheet: string, index: number): string {
  const resolved = resolveRange(workbook, series.nameReference, fallbackSheet);
  if (resolved) {
    const text = labelOf(resolved.sheet.getValue(resolved.range.top, resolved.range.left)).trim();
    if (text.length > 0) return text;
  }
  if (series.name?.trim()) return series.name.trim();
  return `Series ${index}`;
}

/** Names made unique for the legend: a repeat gets " (2)", " (3)"… */
function uniqueNames(names: readonly string[]): string[] {
  const seen = new Map<string, number>();
  return names.map((name) => {
    const n = (seen.get(name) ?? 0) + 1;
    seen.set(name, n);
    return n === 1 ? name : `${name} (${n})`;
  });
}

const gapAt = (list: readonly number[] | null, i: number): boolean => Number.isNaN(list?.[i] ?? NaN);

/**
 * An X/Y series without its fully empty rows (no Y, X or size). A series without X values plots
 * against the row number, so only its trailing empty rows go.
 */
function dropEmptyPoints(series: ResolvedSeries): ResolvedSeries {
  const { values, xValues, bubbleSizes } = series;
  const empty = (i: number) => gapAt(values, i) && gapAt(xValues, i) && gapAt(bubbleSizes, i);
  const length = Math.max(values.length, xValues?.length ?? 0, bubbleSizes?.length ?? 0);
  let keep: number[];
  if (xValues) {
    keep = [];
    for (let i = 0; i < length; i++) if (!empty(i)) keep.push(i);
  } else {
    let end = length;
    while (end > 0 && empty(end - 1)) end--;
    keep = Array.from({ length: end }, (_, i) => i);
  }
  if (keep.length === length) return series;
  const pick = (list: readonly number[]) => keep.map((i) => list[i] ?? NaN);
  return {
    ...series,
    values: pick(values),
    xValues: xValues ? pick(xValues) : null,
    bubbleSizes: bubbleSizes ? pick(bubbleSizes) : null,
  };
}

/**
 * Resolves a chart definition's stored range strings (e.g. `Data.B2:B7`) against a workbook into
 * numbers and labels. Unqualified references default to the first data sheet. A chart bound to a
 * table reads the table's current columns instead (`ChartTableBinding`). Rows that repeat a category
 * label merge into one category by the chart's aggregate (`ChartGrouping`); scatter and bubble
 * charts plot every row that has anything in it. Series names are unique, so the legend never
 * lists one twice.
 */
export const ChartDataResolver = {
  resolve(workbook: Workbook, definition: ChartDefinition): ResolvedChart {
    const chart = ChartTableBinding.materialize(workbook, definition);
    const fallbackSheet = defaultSheetName(workbook);
    const labels = readLabels(workbook, chart.categoriesReference, fallbackSheet);
    const cells = chart.series.map((s) => readCells(workbook, s.valuesReference, fallbackSheet));
    // Rows that share a category label become one category; with no labels every row is its own.
    const grouped =
      ChartTypes.groupsCategories(chart.type) && labels.length > 0
        ? ChartGrouping.group(labels, cells, chart.aggregate)
        : { categories: labels, values: cells.map((c) => c.values) };
    const categories = grouped.categories;
    const xy = !ChartTypes.groupsCategories(chart.type);
    const names = uniqueNames(chart.series.map((s, i) => seriesName(workbook, s, fallbackSheet, i + 1)));
    const series: ResolvedSeries[] = chart.series.map((definition, i) => {
      const xValues = readNumbers(workbook, definition.xValuesReference, fallbackSheet);
      const bubbleSizes = readNumbers(workbook, definition.bubbleSizeReference, fallbackSheet);
      const resolved: ResolvedSeries = {
        name: names[i]!,
        type: definition.type ?? chart.type,
        values: grouped.values[i]!,
        xValues: xValues.length > 0 ? xValues : null,
        bubbleSizes: bubbleSizes.length > 0 ? bubbleSizes : null,
      };
      return xy ? dropEmptyPoints(resolved) : resolved;
    });
    return {
      title: chart.title,
      type: chart.type,
      stacking: chart.stacking,
      showLegend: chart.showLegend,
      xAxisTitle: chart.xAxisTitle,
      yAxisTitle: chart.yAxisTitle,
      categories,
      series,
      aggregate: chart.aggregate,
      rowCounts: grouped.rowCounts ?? null,
      isEmpty: series.length === 0 || series.every((s) => s.values.length === 0),
    };
  },
} as const;
