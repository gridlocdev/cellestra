import {
  ChartAggregates,
  ChartTypes,
  createChart,
  createSeries,
  type ChartAggregate,
  type ChartDefinition,
  type ChartStacking,
  type ChartType,
} from '@cellestra/charts';
import {
  CellRange,
  ChartTableBinding,
  findTableByName,
  seedChartFromRange,
  type Workbook,
  type Worksheet,
} from '@cellestra/core';
import { pickSheet, resolveRange } from '../infra/cells.ts';
import type { EditSession } from '../infra/editSession.ts';
import { CliError } from '../infra/errors.ts';

/**
 * Chart sheets from the command line — the headless face of Insert ▸ Chart Sheet. A chart is seeded
 * from a tabular range (`from`) or wired field by field (`values`, `categories`, `xValues`,
 * `bubbleSizes`, `names`); each type accepts only the fields it can plot, checked up front. A
 * `from` that names a table binds the chart to it, so it follows the table as it grows.
 */

export interface ChartData {
  readonly values: readonly string[];
  readonly names: readonly string[];
  readonly categories: string | null;
  readonly xValues: string | null;
  readonly bubbleSizes: string | null;
  readonly from: string | null;
}

export interface ChartLook {
  readonly title: string | null;
  readonly xTitle: string | null;
  readonly yTitle: string | null;
  readonly stacking: string | null;
  readonly aggregate: string | null;
}

/** A chart type as the CLI prints it, all lowercase: `filledRadar` → `filledradar`. */
const spellType = (type: ChartType): string => type.toLowerCase();

/** Build, validate, and add one chart sheet — shared by the `chart` verb and `apply`. */
export function addChart(
  session: EditSession,
  typeText: string,
  look: ChartLook,
  noLegend: boolean,
  data: ChartData,
): void {
  const definition = buildDefinition(session.workbook, typeText, look, noLegend, data);
  const sheet = session.workbook.addChartSheet(definition);
  session.field('sheet', sheet.name);
  session.note(`Added chart sheet '${sheet.name}'.`);
}

/**
 * Change an existing chart in place. Only the options given change; giving data (`from` or
 * `values`) rebuilds the series, otherwise they stay. The sheet keeps its name.
 */
export function updateChart(
  session: EditSession,
  sheetName: string,
  typeText: string | null,
  look: ChartLook,
  legend: boolean | null,
  data: ChartData,
): void {
  const sheet = chartSheet(session.workbook, sheetName);
  const current = sheet.chart!;
  const seeded = data.from !== null && data.from.length > 0;
  const rebuildData = data.values.length > 0 || seeded;
  const chartType = typeText === null ? current.type : parseType(typeText);

  let next: ChartDefinition;
  if (rebuildData) {
    // Only the data given is rebuilt: the parts the chart already had carry over, so `--values`
    // alone doesn't quietly drop the categories (or a scatter's shared X range).
    const first = current.series[0];
    next = buildDefinition(
      session.workbook,
      typeText ?? current.type,
      {
        title: look.title ?? current.title,
        xTitle: look.xTitle ?? current.xAxisTitle,
        yTitle: look.yTitle ?? current.yAxisTitle,
        stacking: look.stacking ?? current.stacking,
        aggregate: look.aggregate ?? keptAggregate(current, chartType),
      },
      legend === false || (legend === null && !current.showLegend),
      {
        values: data.values,
        names: data.names,
        categories: data.categories ?? (seeded ? null : current.categoriesReference),
        xValues: data.xValues ?? (seeded ? null : (first?.xValuesReference ?? null)),
        bubbleSizes: data.bubbleSizes ?? (seeded ? null : (first?.bubbleSizeReference ?? null)),
        from: data.from,
      },
    );
  } else {
    // A bound chart keeps its columns where the new type can plot them; a typed range replaces
    // the binding with plain ranges.
    next = ChartTableBinding.retype(session.workbook, current, chartType);
    if (data.categories !== null && next.source) next = ChartTableBinding.toRanges(session.workbook, next);
    next.title = look.title ?? current.title;
    next.showLegend = legend ?? current.showLegend;
    next.xAxisTitle = look.xTitle ?? current.xAxisTitle;
    next.yAxisTitle = look.yTitle ?? current.yAxisTitle;
    next.stacking = parseStacking(look.stacking ?? current.stacking, chartType);
    next.aggregate = parseAggregate(look.aggregate ?? keptAggregate(current, chartType), chartType);
    if (data.categories !== null) next.categoriesReference = requireRange(data.categories, '--categories');
    validate(next, chartType);
  }

  next.objectName = current.objectName; // the same embedded object, rewritten
  sheet.chart = next;
  session.field('sheet', sheet.name);
  session.note(`Updated chart sheet '${sheet.name}'.`);
}

export function chartSheet(workbook: Workbook, name: string): Worksheet {
  const sheet = pickSheet(workbook, name);
  if (!sheet.isChartSheet) {
    throw new CliError(`'${sheet.name}' is not a chart sheet. \`cellestra charts list\` shows them.`);
  }
  return sheet;
}

/** Everything a chart is, as JSON — what `charts show` prints and `charts list` summarizes. */
export function describeChart(sheet: Worksheet): Record<string, unknown> {
  const c = sheet.chart!;
  return {
    sheet: sheet.name,
    type: spellType(c.type),
    title: c.title,
    stacking: c.stacking === 'none' ? null : c.stacking,
    aggregate: c.aggregate === 'sum' ? null : c.aggregate,
    legend: c.showLegend,
    xTitle: c.xAxisTitle,
    yTitle: c.yAxisTitle,
    table: c.source && {
      name: c.source.table,
      categories: c.source.categoryColumn,
      columns: c.source.series.map((s) => s.column),
      sizeColumns: c.type === 'bubble' ? c.source.series.map((s) => s.sizeColumn) : undefined,
    },
    categories: c.categoriesReference,
    series: c.series.map((s) => ({
      name: s.name,
      nameRef: s.nameReference,
      values: s.valuesReference,
      xValues: s.xValuesReference,
      bubbleSizes: s.bubbleSizeReference,
      type: s.type === null ? null : spellType(s.type),
    })),
  };
}

export const chartTypeName = spellType;

function buildDefinition(
  workbook: Workbook,
  typeText: string,
  look: ChartLook,
  noLegend: boolean,
  data: ChartData,
): ChartDefinition {
  const chartType = parseType(typeText);
  const definition = createChart({
    type: chartType,
    title: look.title ?? 'Chart',
    showLegend: !noLegend,
    xAxisTitle: look.xTitle,
    yAxisTitle: look.yTitle,
    stacking: parseStacking(look.stacking, chartType),
    aggregate: parseAggregate(look.aggregate, chartType),
  });
  let built = definition;
  if (data.values.length > 0) buildExplicit(definition, chartType, data);
  else if (data.from !== null && data.from.length > 0)
    built = buildFrom(definition, chartType, workbook, data.from);
  else throw new CliError('Give the data to plot with --from <range|table> or one or more --values <range>.');
  validate(built, chartType);
  return built;
}

const isBlank = (s: string | null): boolean => s === null || s.trim().length === 0;

function buildExplicit(definition: ChartDefinition, type: ChartType, data: ChartData): void {
  const xy = ChartTypes.isXY(type);
  if (xy && isBlank(data.xValues)) throw new CliError(`A ${spellType(type)} chart needs --x-values.`);
  if (type === 'bubble' && isBlank(data.bubbleSizes))
    throw new CliError('A bubble chart needs --bubble-sizes.');
  if (!xy && !isBlank(data.categories)) {
    definition.categoriesReference = requireRange(data.categories!, '--categories');
  }
  const xRef = xy ? requireRange(data.xValues!, '--x-values') : null;
  const sizeRef = type === 'bubble' ? requireRange(data.bubbleSizes!, '--bubble-sizes') : null;
  data.values.forEach((values, i) => {
    definition.series.push(
      createSeries({
        valuesReference: requireRange(values, '--values'),
        xValuesReference: xRef,
        bubbleSizeReference: sizeRef,
        name: data.names[i] ?? null,
      }),
    );
  });
}

/**
 * Seed the chart from `from`, the way Insert ▸ Chart Sheet seeds it from a selection: a table name
 * binds the chart to the table (text column as categories, number columns as series); a range is
 * read as a tabular block.
 */
function buildFrom(
  definition: ChartDefinition,
  type: ChartType,
  workbook: Workbook,
  from: string,
): ChartDefinition {
  const table = findTableByName(workbook, from.trim());
  if (table) {
    const source = ChartTableBinding.seed(workbook, table.table.name, type)!;
    if (source.series.length === 0 || (source.categoryColumn === null && ChartTypes.isXY(type)))
      throw new CliError(
        `The table '${table.table.name}' has too few number columns for a ${spellType(type)} chart.`,
      );
    if (type === 'bubble' && source.series[0]!.sizeColumn === null)
      throw new CliError(
        `A bubble chart needs three number columns (X, Y and size); '${table.table.name}' has fewer.`,
      );
    return ChartTableBinding.bind(workbook, definition, source);
  }
  if (type === 'bubble') {
    throw new CliError(
      'A bubble chart needs three ranges per series; build it with --values, --x-values, and --bubble-sizes, or --from a table.',
    );
  }
  const { sheet, range } = resolveRange(workbook, from);
  if (range.rowCount < 2) {
    throw new CliError('--from needs at least two rows (a header or first data point plus values).');
  }
  seedChartFromRange(definition, sheet, range);
  return definition;
}

function validate(definition: ChartDefinition, type: ChartType): void {
  if (definition.series.length === 0) throw new CliError('The chart has no data series to plot.');
  if ((type === 'pie' || type === 'doughnut') && definition.series.length > 1) {
    throw new CliError(
      `A ${spellType(type)} chart plots a single series; pass one --values (or a two-column --from).`,
    );
  }
}

function parseType(text: string): ChartType {
  switch (text.trim().toLowerCase()) {
    case 'column':
      return 'column';
    case 'bar':
      return 'bar';
    case 'line':
      return 'line';
    case 'area':
      return 'area';
    case 'pie':
      return 'pie';
    case 'doughnut':
    case 'donut':
      return 'doughnut';
    case 'scatter':
      return 'scatter';
    case 'bubble':
      return 'bubble';
    case 'radar':
      return 'radar';
    case 'filled-radar':
    case 'filledradar':
      return 'filledRadar';
    default:
      throw new CliError(
        `Unknown chart type '${text}'. Use one of: column, bar, line, area, pie, doughnut, scatter, bubble, radar, filled-radar.`,
      );
  }
}

function parseStacking(text: string | null, type: ChartType): ChartStacking {
  if (text === null || text.trim().length === 0) return 'none';
  let stacking: ChartStacking;
  switch (text.trim().toLowerCase()) {
    case 'none':
      stacking = 'none';
      break;
    case 'stacked':
      stacking = 'stacked';
      break;
    case 'percent':
    case 'percentage':
      stacking = 'percent';
      break;
    default:
      throw new CliError(`Unknown stacking '${text}'. Use none, stacked, or percent.`);
  }
  if (stacking !== 'none' && !ChartTypes.supportsStacking(type)) {
    throw new CliError(`--stacking applies only to column, bar, and area charts, not ${spellType(type)}.`);
  }
  return stacking;
}

/** The chart's aggregate, kept across a type change when the new type still groups categories. */
const keptAggregate = (current: ChartDefinition, type: ChartType): ChartAggregate =>
  ChartTypes.groupsCategories(type) ? current.aggregate : 'sum';

function parseAggregate(text: string | null, type: ChartType): ChartAggregate {
  if (text === null || text.trim().length === 0) return 'sum';
  const aggregate = ChartAggregates.find((a) => a === text.trim().toLowerCase());
  if (!aggregate) throw new CliError(`Unknown aggregate '${text}'. Use sum, average, count, min, or max.`);
  if (aggregate !== 'sum' && !ChartTypes.groupsCategories(type))
    throw new CliError(`--aggregate applies only to charts with categories, not ${spellType(type)}.`);
  return aggregate;
}

/** Check a chart data reference parses as a range, and return it verbatim for storage. */
function requireRange(reference: string, option: string): string {
  if (!CellRange.tryParse(reference)) {
    throw new CliError(`'${reference}' is not a valid range for ${option} (e.g. Sheet1.B2:B7).`);
  }
  return reference;
}
