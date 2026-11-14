/** Chart model — standalone, no dependency on the spreadsheet engine. */

export type ChartType =
  'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'bubble' | 'radar' | 'filledRadar';

export type ChartStacking = 'none' | 'stacked' | 'percent';

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

export function cloneChart(c: ChartDefinition): ChartDefinition {
  return { ...c, series: c.series.map((s) => ({ ...s })) };
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
} as const;
