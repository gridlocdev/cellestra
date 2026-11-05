import {
  ChartTypes,
  createChart,
  createSeries,
  type ChartDefinition,
  type ChartSeries,
  type ChartType,
} from '../model.ts';
import { ChartXmlNames as N } from './chartXmlNames.ts';

/**
 * The read-only view of a parsed XML element that `ChartDocument.parse` needs. The charts package
 * has no XML parser of its own (it must not depend on the engine); the engine's namespace-aware
 * `XmlElement` satisfies this interface structurally, and so would any other tree with the same
 * namespace + local-name lookups.
 */
export interface ChartXmlElement {
  readonly uri: string;
  readonly local: string;
  /** The attribute value, or null when absent. */
  attr(uri: string, local: string): string | null;
  /** Concatenated descendant text. */
  readonly value: string;
  readonly nodes: readonly unknown[];
}

/**
 * A from-scratch reader/writer for an embedded ODF chart object's `content.xml` — the
 * `<office:chart><chart:chart>` tree. The model maps exactly onto the ODF the spec calls for:
 * - the chart family is the `chart:class` on `<chart:chart>`;
 * - column-vs-bar is the `chart:vertical` flag on the plot area;
 * - stacking / 100%-stacking are the `chart:stacked` / `chart:percentage` flags;
 * - a combo is per-series `chart:class` on each `<chart:series>`.
 */
export const ChartDocument = {
  /** The default frame size for a chart sheet (the whole "page"), in centimetres. */
  defaultWidthCm: 24.0,
  defaultHeightCm: 14.0,

  /** Build the `content.xml` text for an embedded chart object. */
  buildContent(chart: ChartDefinition): string {
    const supportsStacking = ChartTypes.supportsStacking(chart.type);
    const stacked = (chart.stacking === 'stacked' || chart.stacking === 'percent') && supportsStacking;
    const percentage = chart.stacking === 'percent' && supportsStacking;

    const plotArea = el('chart:plot-area');
    if (stacked) plotArea.attrs.push(['chart:stacked', 'true']);
    if (percentage) plotArea.attrs.push(['chart:percentage', 'true']);
    // ODF: bars run horizontally when the plot area is "vertical".
    if (chart.type === 'bar') plotArea.attrs.push(['chart:vertical', 'true']);

    // X / category axis (carries the categories range, ODF §18).
    const xAxis = el('chart:axis', [
      ['chart:dimension', 'x'],
      ['chart:name', 'primary-x'],
    ]);
    if (chart.xAxisTitle) xAxis.children.push(title(chart.xAxisTitle));
    if (chart.categoriesReference && chart.categoriesReference.trim()) {
      xAxis.children.push(el('chart:categories', [['table:cell-range-address', chart.categoriesReference]]));
    }
    plotArea.children.push(xAxis);

    // Y / value axis.
    const yAxis = el('chart:axis', [
      ['chart:dimension', 'y'],
      ['chart:name', 'primary-y'],
    ]);
    if (chart.yAxisTitle) yAxis.children.push(title(chart.yAxisTitle));
    plotArea.children.push(yAxis);

    for (const series of chart.series) plotArea.children.push(buildSeries(series, chart.type));

    const chartElement = el('chart:chart', [
      ['chart:class', ChartTypes.odsClass(chart.type)],
      ['svg:width', cm(ChartDocument.defaultWidthCm)],
      ['svg:height', cm(ChartDocument.defaultHeightCm)],
    ]);
    if (chart.title) chartElement.children.push(title(chart.title));
    if (chart.showLegend) chartElement.children.push(el('chart:legend', [['chart:legend-position', 'end']]));
    chartElement.children.push(plotArea);

    const root = el(
      'office:document-content',
      [
        ['xmlns:office', N.office],
        ['xmlns:chart', N.chart],
        ['xmlns:table', N.table],
        ['xmlns:text', N.text],
        ['xmlns:svg', N.svg],
        ['office:version', '1.3'],
      ],
      [el('office:body', [], [el('office:chart', [], [chartElement])])],
    );

    return '<?xml version="1.0" encoding="UTF-8"?>\n' + write(root, 0);
  },

  /**
   * Parse an embedded chart object's `content.xml` root element back into a definition. Null when
   * there is no document, or it isn't an `office:chart` document.
   */
  parse(root: ChartXmlElement | null | undefined): ChartDefinition | null {
    if (!root) return null;
    const chartElement = child(child(child(root, N.office, 'body'), N.office, 'chart'), N.chart, 'chart');
    if (!chartElement) return null;

    const plotArea = child(chartElement, N.chart, 'plot-area');
    const vertical = plotArea?.attr(N.chart, 'vertical') === 'true';
    const type = ChartTypes.fromOds(chartElement.attr(N.chart, 'class'), vertical);
    const stacked = plotArea?.attr(N.chart, 'stacked') === 'true';
    const percentage = plotArea?.attr(N.chart, 'percentage') === 'true';

    const chart = createChart({
      type,
      title: titleText(chartElement) ?? '',
      showLegend: child(chartElement, N.chart, 'legend') !== null,
      stacking: percentage ? 'percent' : stacked ? 'stacked' : 'none',
    });
    if (!plotArea) return chart;

    for (const axis of elements(plotArea, N.chart, 'axis')) {
      const dimension = axis.attr(N.chart, 'dimension');
      const axisTitle = titleText(axis);
      if (dimension === 'x') {
        chart.xAxisTitle = axisTitle;
        const categories = child(axis, N.chart, 'categories')?.attr(N.table, 'cell-range-address');
        if (categories && categories.trim()) chart.categoriesReference = categories;
      } else if (dimension === 'y') {
        chart.yAxisTitle = axisTitle;
      }
    }

    for (const seriesElement of elements(plotArea, N.chart, 'series')) {
      const seriesType = ChartTypes.fromOds(seriesElement.attr(N.chart, 'class'), vertical);
      chart.series.push(
        createSeries({
          valuesReference: seriesElement.attr(N.chart, 'values-cell-range-address') ?? '',
          nameReference: seriesElement.attr(N.chart, 'label-cell-address'),
          name: seriesElement.attr(N.cellestraChart, 'name'),
          xValuesReference:
            child(seriesElement, N.chart, 'domain')?.attr(N.table, 'cell-range-address') ?? null,
          type: seriesType === type ? null : seriesType,
        }),
      );
    }
    return chart;
  },
} as const;

function buildSeries(series: ChartSeries, chartType: ChartType): Node {
  const type = series.type ?? chartType;
  const element = el('chart:series', [
    ['chart:class', ChartTypes.odsClass(type)],
    ['chart:values-cell-range-address', series.valuesReference],
  ]);
  if (series.nameReference && series.nameReference.trim()) {
    element.attrs.push(['chart:label-cell-address', series.nameReference]);
  } else if (series.name) {
    // ODF has no attribute for a literal (non-cell) series name, so keep it on a Cellestra-namespaced
    // attribute. Conforming readers ignore unknown-namespace attributes, so this round-trips for us
    // without confusing LibreOffice.
    element.attrs.push(['xmlns:cellestra', N.cellestraChart], ['cellestra:name', series.name]);
  }
  if (series.xValuesReference && series.xValuesReference.trim()) {
    element.children.push(el('chart:domain', [['table:cell-range-address', series.xValuesReference]]));
  }
  return element;
}

function title(text: string): Node {
  return el('chart:title', [], [el('text:p', [], [text])]);
}

function titleText(parent: ChartXmlElement): string | null {
  const text = child(child(parent, N.chart, 'title'), N.text, 'p')?.value;
  return text ? text : null;
}

function cm(value: number): string {
  return `${Number.parseFloat(value.toFixed(3))}cm`;
}

// ---- Reading through the structural element view -------------------------

function isElement(n: unknown): n is ChartXmlElement {
  return typeof n === 'object' && n !== null && 'local' in n && 'uri' in n && 'attr' in n;
}

function elements(parent: ChartXmlElement | null, uri: string, local: string): ChartXmlElement[] {
  if (!parent) return [];
  return parent.nodes.filter((n): n is ChartXmlElement => isElement(n) && n.uri === uri && n.local === local);
}

function child(parent: ChartXmlElement | null, uri: string, local: string): ChartXmlElement | null {
  return elements(parent, uri, local)[0] ?? null;
}

// ---- A tiny writer (the chart part only uses fixed, known prefixes) -------

interface Node {
  readonly q: string;
  readonly attrs: Array<[string, string]>;
  readonly children: Array<Node | string>;
}

function el(q: string, attrs: Array<[string, string]> = [], children: Array<Node | string> = []): Node {
  return { q, attrs, children };
}

function write(n: Node, depth: number): string {
  let s = `<${n.q}`;
  for (const [k, v] of n.attrs) s += ` ${k}="${escapeAttr(v)}"`;
  if (n.children.length === 0) return s + '/>';
  s += '>';
  const mixed = n.children.some((c) => typeof c === 'string');
  for (const c of n.children) {
    if (typeof c === 'string') s += escapeText(c);
    else s += (mixed ? '' : '\n' + '  '.repeat(depth + 1)) + write(c, depth + 1);
  }
  if (!mixed) s += '\n' + '  '.repeat(depth);
  return s + `</${n.q}>`;
}

/**
 * Characters XML 1.0 cannot carry at all, escaped or not: C0 controls other than tab, newline and
 * carriage return, and U+FFFE/U+FFFF. One pasted into a chart title made object/content.xml
 * unparseable, and LibreOffice dropped the whole chart. (Duplicated from core's XML writer, which
 * this package cannot import.)
 */
// eslint-disable-next-line no-control-regex
const INVALID_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;

function escapeText(s: string): string {
  return s
    .replace(INVALID_XML_CHARS, '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('\r', '&#xD;');
}

function escapeAttr(s: string): string {
  return escapeText(s).replaceAll('"', '&quot;').replaceAll('\n', '&#xA;').replaceAll('\t', '&#x9;');
}
