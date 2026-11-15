import type { ChartDefinition, ChartSeries, ChartStacking, ChartType } from '@cellestra/charts';
import { Defaults } from '../defaults.ts';
import { fixedDigits } from '../formulas/evaluation/decimalDigits.ts';
import { quoteSheetIfNeeded } from '../model/cellAddress.ts';
import { CellRange } from '../model/cellRange.ts';
import { ChartTableBinding } from './chartTableBinding.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import type { CellValue } from '../values/cellValue.ts';

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
  /** True when there is nothing to draw (no series, or every series is empty). */
  readonly isEmpty: boolean;
}

/** .NET `ToString("0.######", InvariantCulture)`. */
function invariantLabel(n: number): string {
  if (!Number.isFinite(n)) return Number.isNaN(n) ? 'NaN' : n > 0 ? 'Infinity' : '-Infinity';
  const { integer, fraction } = fixedDigits(Math.abs(n), 6);
  const trimmed = fraction.replace(/0+$/, '');
  const body = integer + (trimmed ? '.' + trimmed : '');
  return n < 0 && /[1-9]/.test(body) ? '-' + body : body;
}

function labelOf(value: CellValue): string {
  switch (value.kind) {
    case 'text':
      return value.value;
    case 'number':
      return invariantLabel(value.value);
    case 'boolean':
      return value.value ? 'TRUE' : 'FALSE';
    default:
      return '';
  }
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

function readNumbers(workbook: Workbook, reference: string | null, fallbackSheet: string): number[] {
  const resolved = resolveRange(workbook, reference, fallbackSheet);
  if (!resolved) return [];
  const values: number[] = [];
  for (const [row, column] of resolved.range.positions()) {
    const v = resolved.sheet.getValue(row, column);
    values.push(v.kind === 'number' ? v.value : v.kind === 'boolean' ? (v.value ? 1 : 0) : NaN);
  }
  return values;
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
    const text = labelOf(resolved.sheet.getValue(resolved.range.top, resolved.range.left));
    if (text.length > 0) return text;
  }
  if (series.name) return series.name;
  return `Series ${index}`;
}

/**
 * Resolves a chart definition's stored range strings (e.g. `Data.B2:B7`) against a workbook into
 * numbers and labels. Unqualified references default to the first data sheet. A chart bound to a
 * table reads the table's current columns instead (`ChartTableBinding`).
 */
export const ChartDataResolver = {
  resolve(workbook: Workbook, definition: ChartDefinition): ResolvedChart {
    const chart = ChartTableBinding.materialize(workbook, definition);
    const fallbackSheet = defaultSheetName(workbook);
    const categories = readLabels(workbook, chart.categoriesReference, fallbackSheet);
    const series: ResolvedSeries[] = chart.series.map((definition, i) => {
      const xValues = readNumbers(workbook, definition.xValuesReference, fallbackSheet);
      const bubbleSizes = readNumbers(workbook, definition.bubbleSizeReference, fallbackSheet);
      return {
        name: seriesName(workbook, definition, fallbackSheet, i + 1),
        type: definition.type ?? chart.type,
        values: readNumbers(workbook, definition.valuesReference, fallbackSheet),
        xValues: xValues.length > 0 ? xValues : null,
        bubbleSizes: bubbleSizes.length > 0 ? bubbleSizes : null,
      };
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
      isEmpty: series.length === 0 || series.every((s) => s.values.length === 0),
    };
  },
} as const;
