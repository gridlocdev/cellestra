import type { ChartDefinition, ChartSeries } from '@cellestra/charts';
import { cloneChart, createSeries } from '@cellestra/charts';
import { quoteSheetIfNeeded } from '../model/cellAddress.ts';
import { CellRange } from '../model/cellRange.ts';
import type { Workbook } from '../model/workbook.ts';

/** The cell-range fields of a chart a reference can be picked for on the sheet. */
export type ChartSeriesFieldKind = 'values' | 'xValues' | 'bubbleSizes' | 'nameCell';

/** One reference field of a chart: its categories, or one field of series `series` (zero-based). */
export type ChartField =
  { readonly kind: 'categories' } | { readonly kind: ChartSeriesFieldKind; readonly series: number };

const SERIES_LABELS: Record<ChartSeriesFieldKind, string> = {
  values: 'values',
  xValues: 'X values',
  bubbleSizes: 'bubble sizes',
  nameCell: 'name cell',
};

function seriesReference(series: ChartSeries | undefined, kind: ChartSeriesFieldKind): string | null {
  if (!series) return null;
  switch (kind) {
    case 'values':
      return series.valuesReference;
    case 'xValues':
      return series.xValuesReference;
    case 'bubbleSizes':
      return series.bubbleSizeReference;
    case 'nameCell':
      return series.nameReference;
  }
}

/**
 * A chart's reference fields by name, for the panel's "Pick on sheet": read one, write one, and
 * know which sheet to show while picking.
 */
export const ChartFields = {
  /** The reference in `field`, or null when it is empty or the series doesn't exist. */
  get(chart: ChartDefinition, field: ChartField): string | null {
    const text =
      field.kind === 'categories'
        ? chart.categoriesReference
        : seriesReference(chart.series[field.series], field.kind);
    return text && text.trim() !== '' ? text : null;
  },

  /** The chart with `field` set to `reference`; a series one past the last is added. */
  set(chart: ChartDefinition, field: ChartField, reference: string): ChartDefinition {
    const next = cloneChart(chart);
    if (field.kind === 'categories') {
      next.categoriesReference = reference;
      return next;
    }
    if (field.series === next.series.length) next.series.push(createSeries({ valuesReference: '' }));
    const series = next.series[field.series];
    if (!series) return next;
    switch (field.kind) {
      case 'values':
        series.valuesReference = reference;
        break;
      case 'xValues':
        series.xValuesReference = reference;
        break;
      case 'bubbleSizes':
        series.bubbleSizeReference = reference;
        break;
      case 'nameCell':
        series.nameReference = reference;
        break;
    }
    return next;
  },

  /** How the panel and the pick prompt name `field`: “the category labels”, “series 2 values”. */
  label(field: ChartField): string {
    return field.kind === 'categories'
      ? 'the category labels'
      : `series ${field.series + 1} ${SERIES_LABELS[field.kind]}`;
  },

  /** `Sheet.A1:B4`, the way a chart stores a reference. */
  reference(sheetName: string, range: CellRange): string {
    const plain = CellRange.fromBounds(range.top, range.left, range.bottom, range.right);
    return `${quoteSheetIfNeeded(sheetName)}.${plain.toString()}`;
  },

  /**
   * Where to pick `field`: the sheet and range it already points at; else the sheet of any other
   * reference in the chart; else the first sheet that isn't a chart sheet. Null when there is none.
   */
  pickTarget(
    workbook: Workbook,
    chart: ChartDefinition,
    field: ChartField,
  ): { sheet: string; range: CellRange | null } | null {
    const fallback = workbook.worksheets.find((s) => !s.isChartSheet)?.name ?? null;
    const parse = (text: string | null) => {
      const range = text ? CellRange.tryParse(text.trim()) : null;
      if (!range) return null;
      const sheet = workbook.findSheet(range.sheet ?? fallback ?? '');
      return sheet && !sheet.isChartSheet ? { sheet: sheet.name, range: range.withSheet(null) } : null;
    };
    const own = parse(ChartFields.get(chart, field));
    if (own) return own;
    const others = [
      chart.categoriesReference,
      ...chart.series.flatMap((s) => [s.valuesReference, s.xValuesReference, s.bubbleSizeReference]),
    ];
    for (const text of others) {
      const other = parse(text);
      if (other) return { sheet: other.sheet, range: null };
    }
    return fallback === null ? null : { sheet: fallback, range: null };
  },
} as const;
