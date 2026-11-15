/** Chart model — standalone, no dependency on the spreadsheet engine. */

export type ChartType =
  'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'bubble' | 'radar' | 'filledRadar';

export type ChartStacking = 'none' | 'stacked' | 'percent';

/** How rows that share a category label combine into one category. */
export type ChartAggregate = 'sum' | 'average' | 'count' | 'min' | 'max';

export interface ChartSeries {
  name: string | null;
  /** A cell holding the series name (overrides `name` when resolvable). */
  nameReference: string | null;
  valuesReference: string;
  xValuesReference: string | null;
  bubbleSizeReference: string | null;
  /** Per-series override (combo charts). */
  type: ChartType | null;
}

/** One plotted column of a table-bound chart, by its header name. */
export interface ChartSourceSeries {
  /** The header of the column holding the values (Y values for scatter and bubble). */
  column: string;
  /** Per-series override (combo charts), as `ChartSeries.type`. */
  type: ChartType | null;
  /** Bubble charts: the header of the column sizing each bubble. */
  sizeColumn: string | null;
}

/**
 * A chart bound to a named table by column names rather than cell ranges, so it follows the table
 * as it grows and its columns move. The engine turns it into ranges (`ChartTableBinding`); the
 * ranges are what the file stores for other readers.
 */
export interface ChartTableSource {
  /** The table's name (matched without case). */
  table: string;
  /** The header of the column labelling the categories, or holding the X values of an X/Y chart. */
  categoryColumn: string | null;
  series: ChartSourceSeries[];
}

export interface ChartDefinition {
  title: string;
  type: ChartType;
  stacking: ChartStacking;
  categoriesReference: string | null;
  series: ChartSeries[];
  showLegend: boolean;
  xAxisTitle: string | null;
  yAxisTitle: string | null;
  /** The embedded object's directory inside the .ods (`Object 1`). */
  objectName: string;
  /** The table the data comes from, by column names; null plots the ranges above as they are. */
  source: ChartTableSource | null;
  /**
   * How rows with the same category label combine (category charts only; scatter and bubble plot
   * every row).
   */
  aggregate: ChartAggregate;
}

export function createChart(init: Partial<ChartDefinition> = {}): ChartDefinition {
  return {
    title: '',
    type: 'column',
    stacking: 'none',
    categoriesReference: null,
    series: [],
    showLegend: true,
    xAxisTitle: null,
    yAxisTitle: null,
    objectName: '',
    source: null,
    aggregate: 'sum',
    ...init,
  };
}

export function createSeries(init: Partial<ChartSeries> & Pick<ChartSeries, 'valuesReference'>): ChartSeries {
  return {
    name: null,
    nameReference: null,
    xValuesReference: null,
    bubbleSizeReference: null,
    type: null,
    ...init,
  };
}

export function cloneSource(source: ChartTableSource): ChartTableSource {
  return { ...source, series: source.series.map((s) => ({ ...s })) };
}

export function cloneChart(c: ChartDefinition): ChartDefinition {
  return { ...c, series: c.series.map((s) => ({ ...s })), source: c.source ? cloneSource(c.source) : null };
}

const ODS_CLASS: Record<ChartType, string> = {
  column: 'chart:bar',
  bar: 'chart:bar',
  line: 'chart:line',
  area: 'chart:area',
  pie: 'chart:circle',
  doughnut: 'chart:ring',
  scatter: 'chart:scatter',
  bubble: 'chart:bubble',
  radar: 'chart:radar',
  filledRadar: 'chart:filled-radar',
};

export const ChartTypes = {
  all: Object.keys(ODS_CLASS) as ChartType[],
  odsClass: (type: ChartType): string => ODS_CLASS[type],
  isHorizontalBars: (type: ChartType): boolean => type === 'bar',
  /** Whether `odsClass` (`chart:bar`, `line`) is a chart family the model can hold. */
  isKnownOds(odsClass: string): boolean {
    const cls = odsClass.includes(':') ? odsClass : `chart:${odsClass}`;
    return Object.values(ODS_CLASS).includes(cls);
  },
  fromOds(odsClass: string | null | undefined, vertical: boolean): ChartType {
    const cls = !odsClass ? '' : odsClass.includes(':') ? odsClass : `chart:${odsClass}`;
    switch (cls) {
      case 'chart:bar':
        return vertical ? 'bar' : 'column';
      case 'chart:line':
        return 'line';
      case 'chart:area':
        return 'area';
      case 'chart:circle':
        return 'pie';
      case 'chart:ring':
        return 'doughnut';
      case 'chart:scatter':
        return 'scatter';
      case 'chart:bubble':
        return 'bubble';
      case 'chart:radar':
        return 'radar';
      case 'chart:filled-radar':
        return 'filledRadar';
      default:
        return 'column';
    }
  },
  usesCategoryAxis: (t: ChartType): boolean =>
    ['column', 'bar', 'line', 'area', 'radar', 'filledRadar'].includes(t),
  supportsStacking: (t: ChartType): boolean => t === 'column' || t === 'bar' || t === 'area',
  isXY: (t: ChartType): boolean => t === 'scatter' || t === 'bubble',
  /** Whether rows sharing a category label merge into one category (every chart but scatter and bubble). */
  groupsCategories: (t: ChartType): boolean => t !== 'scatter' && t !== 'bubble',
} as const;

/** Every aggregate, in the order the chart panel lists them. */
export const ChartAggregates: readonly ChartAggregate[] = ['sum', 'average', 'count', 'min', 'max'];
