import { labelOf } from './chartLabels.ts';
import type { ChartDefinition, ChartSeries, ChartTableSource, ChartType } from '@cellestra/charts';
import { ChartTypes, cloneChart, cloneSource, createSeries } from '@cellestra/charts';
import { DataRegion } from '../data/dataRegion.ts';
import { findTableByName, tableHeaderLabels } from '../data/tables.ts';
import type { Table } from '../data/types.ts';
import { quoteSheetIfNeeded } from '../model/cellAddress.ts';
import { CellRange } from '../model/cellRange.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';

/** What a table column mostly holds, for picking defaults. */
export type TableColumnKind = 'number' | 'text' | 'empty';

/** One column of a table as the chart panel lists it. */
export interface TableColumnInfo {
  /** The header text (or `ColumnN` when the table has no header row). */
  readonly name: string;
  /** The sheet column, zero-based. */
  readonly column: number;
  readonly kind: TableColumnKind;
}

/**
 * Why a chart's ranges can't become a table binding: no data yet, a reference that isn't a cell
 * range, ranges outside every table, ranges that aren't whole table columns, series with different
 * X columns, or no table in the workbook at all.
 */
export type FromRangesFailure =
  'noData' | 'notRanges' | 'notInTable' | 'notWholeColumns' | 'mixedX' | 'noTable';

/** A chart's ranges as a table binding, or why not: a code for the panel and a sentence to show. */
export type FromRangesResult =
  | { readonly ok: true; readonly chart: ChartDefinition }
  | { readonly ok: false; readonly code: FromRangesFailure; readonly reason: string };

const fail = (code: FromRangesFailure, reason: string): FromRangesResult => ({ ok: false, code, reason });

/** Every cell-range reference a chart plots from (not its name cells), skipping blanks. */
export function plottedReferences(chart: ChartDefinition): string[] {
  const xy = ChartTypes.isXY(chart.type);
  return [
    xy ? null : chart.categoriesReference,
    ...chart.series.flatMap((s) => [
      s.valuesReference,
      xy ? s.xValuesReference : null,
      chart.type === 'bubble' ? s.bubbleSizeReference : null,
    ]),
  ].filter((r): r is string => r !== null && r.trim() !== '');
}

/** The sheet an unqualified chart reference points at: the first sheet with cells. */
export function fallbackSheetName(workbook: Workbook): string {
  return workbook.worksheets.find((s) => !s.isChartSheet)?.name ?? '';
}

const letters = (column: number): string => SheetDimensions.columnToLetters(column);

/** The first and last data row of a table (the rows under its header). */
function dataRows(table: Table): { top: number; bottom: number } {
  const top = table.hasHeader ? table.range.top + 1 : table.range.top;
  return { top, bottom: Math.max(top, table.range.bottom) };
}

/** `Sheet.B2:B7`: one table column's data rows. */
function columnRange(sheet: Worksheet, table: Table, column: number): string {
  const { top, bottom } = dataRows(table);
  const c = letters(column);
  return `${quoteSheetIfNeeded(sheet.name)}.${c}${top + 1}:${c}${bottom + 1}`;
}

function headerCell(sheet: Worksheet, table: Table, column: number): string {
  return `${quoteSheetIfNeeded(sheet.name)}.${letters(column)}${table.range.top + 1}`;
}

function kindOf(sheet: Worksheet, table: Table, column: number): TableColumnKind {
  const { top } = dataRows(table);
  let numbers = 0;
  let texts = 0;
  for (let row = top; row <= table.range.bottom; row++) {
    const kind = sheet.getValue(row, column).kind;
    if (kind === 'number') numbers++;
    else if (kind === 'text' || kind === 'boolean') texts++;
  }
  if (numbers === 0 && texts === 0) return 'empty';
  return numbers > texts ? 'number' : 'text';
}

/** The sheet column a stored chart reference starts in, when it is on `sheet` inside `table`. */
function storedColumn(reference: string | null | undefined, sheet: Worksheet, table: Table): number | null {
  if (!reference) return null;
  const range = CellRange.tryParse(reference.trim());
  if (!range) return null;
  if (range.sheet !== null && range.sheet.toLowerCase() !== sheet.name.toLowerCase()) return null;
  const column = range.left;
  return column >= table.range.left && column <= table.range.right ? column : null;
}

/**
 * Find a bound column: by header name (ignoring case), preferring the column the chart last
 * resolved to when two headers share the name; failing that, the column the stored range points at,
 * so a renamed header keeps its place.
 */
function findColumn(
  labels: readonly string[],
  table: Table,
  name: string | null,
  stored: number | null,
): number | null {
  if (name === null) return null;
  const lower = name.toLowerCase();
  const matches: number[] = [];
  labels.forEach((label, i) => {
    if (label.toLowerCase() === lower) matches.push(table.range.left + i);
  });
  if (stored !== null && matches.includes(stored)) return stored;
  if (matches.length > 0) return matches[0]!;
  return stored;
}

/** Whether the chart has any range to plot (bound charts always count as having data). */
function hasData(chart: ChartDefinition): boolean {
  if (chart.source) return true;
  if (chart.categoriesReference?.trim()) return true;
  return chart.series.some(
    (s) => s.valuesReference.trim() !== '' || (s.xValuesReference ?? '').trim() !== '',
  );
}

function located(workbook: Workbook, source: ChartTableSource) {
  const found = findTableByName(workbook, source.table);
  if (!found) return null;
  return { ...found, labels: tableHeaderLabels(found.table, found.sheet) };
}

/**
 * Charts bound to a named table (`ChartDefinition.source`). The binding names the table and its
 * columns; `materialize` turns it into the plain ranges the rest of the engine (and the file) use,
 * from the table's current range and headers. A binding to a table that no longer exists leaves the
 * stored ranges as they are.
 */
export const ChartTableBinding = {
  /** The columns of the table called `name`, or null when there is no such table. */
  columns(workbook: Workbook, name: string): TableColumnInfo[] | null {
    const found = findTableByName(workbook, name);
    if (!found) return null;
    const { sheet, table } = found;
    return tableHeaderLabels(table, sheet).map((label, i) => ({
      name: label,
      column: table.range.left + i,
      kind: kindOf(sheet, table, table.range.left + i),
    }));
  },

  /** Every table in the workbook, by name, in sheet order. */
  tableNames(workbook: Workbook): string[] {
    return workbook.worksheets.flatMap((s) => s.tables.map((t) => t.name));
  },

  /**
   * The chart with its ranges rebuilt from its table binding, and the binding's column names brought
   * up to date with the headers they were found under. The same chart when it has no binding or its
   * table is gone. A bound column that can no longer be found drops out with its series.
   */
  materialize(workbook: Workbook, chart: ChartDefinition): ChartDefinition {
    const source = chart.source;
    if (!source) return chart;
    const where = located(workbook, source);
    if (!where) return chart;
    const { sheet, table, labels } = where;
    const xy = ChartTypes.isXY(chart.type);
    const nameAt = (column: number): string => labels[column - table.range.left]!;

    const storedCategory = xy ? chart.series[0]?.xValuesReference : chart.categoriesReference;
    const categoryColumn = findColumn(
      labels,
      table,
      source.categoryColumn,
      storedColumn(storedCategory, sheet, table),
    );
    const categoryRange = categoryColumn === null ? null : columnRange(sheet, table, categoryColumn);

    const next = cloneChart(chart);
    next.categoriesReference = xy ? null : categoryRange;
    next.series = [];
    const nextSource: ChartTableSource = {
      table: table.name,
      categoryColumn: categoryColumn === null ? null : nameAt(categoryColumn),
      series: [],
    };
    source.series.forEach((bound, i) => {
      const stored = chart.series[i];
      const column = findColumn(
        labels,
        table,
        bound.column,
        storedColumn(stored?.valuesReference, sheet, table),
      );
      if (column === null) return;
      const sizeColumn =
        chart.type === 'bubble'
          ? findColumn(
              labels,
              table,
              bound.sizeColumn,
              storedColumn(stored?.bubbleSizeReference, sheet, table),
            )
          : null;
      next.series.push(
        createSeries({
          valuesReference: columnRange(sheet, table, column),
          nameReference: table.hasHeader ? headerCell(sheet, table, column) : null,
          name: table.hasHeader ? null : nameAt(column),
          xValuesReference: xy ? categoryRange : null,
          bubbleSizeReference: sizeColumn === null ? null : columnRange(sheet, table, sizeColumn),
          type: bound.type,
        }),
      );
      nextSource.series.push({
        column: nameAt(column),
        type: bound.type,
        sizeColumn: sizeColumn === null ? bound.sizeColumn : nameAt(sizeColumn),
      });
    });
    next.source = nextSource;
    return next;
  },

  /**
   * A binding for a new chart of `type` over the table called `name`: the first text column labels
   * the categories and the number columns become series. A scatter or bubble chart takes its X from
   * the first number column (a bubble its sizes from the third); a pie or doughnut plots one column.
   */
  seed(workbook: Workbook, name: string, type: ChartType): ChartTableSource | null {
    const found = findTableByName(workbook, name);
    const columns = ChartTableBinding.columns(workbook, name);
    if (!found || !columns) return null;
    const numbers = columns.filter((c) => c.kind === 'number').map((c) => c.name);
    const text = columns.find((c) => c.kind === 'text')?.name ?? null;
    const source: ChartTableSource = { table: found.table.name, categoryColumn: null, series: [] };
    if (ChartTypes.isXY(type)) {
      const [x, ...rest] = numbers;
      source.categoryColumn = x ?? null;
      if (type === 'bubble') {
        const [y, size] = rest;
        if (y) source.series.push({ column: y, type: null, sizeColumn: size ?? null });
      } else {
        for (const y of rest) source.series.push({ column: y, type: null, sizeColumn: null });
      }
      return source;
    }
    source.categoryColumn = text;
    const plotted = type === 'pie' || type === 'doughnut' ? numbers.slice(0, 1) : numbers;
    for (const column of plotted) source.series.push({ column, type: null, sizeColumn: null });
    return source;
  },

  /** The chart bound to `source`, with its ranges filled in from the table. */
  bind(workbook: Workbook, chart: ChartDefinition, source: ChartTableSource): ChartDefinition {
    return ChartTableBinding.materialize(workbook, { ...chart, source });
  },

  /** Table → Range: the binding's ranges as plain ranges, with no binding left. */
  toRanges(workbook: Workbook, chart: ChartDefinition): ChartDefinition {
    const next = cloneChart(ChartTableBinding.materialize(workbook, chart));
    next.source = null;
    return next;
  },

  /**
   * Range → Table: when every range of the chart is one whole column of one table's data rows, the
   * same chart bound to the table by column names (series take the column headers as their names).
   * Otherwise why not, in a sentence for the panel.
   */
  fromRanges(
    workbook: Workbook,
    chart: ChartDefinition,
    options: { readonly wholeColumns?: boolean } = {},
  ): FromRangesResult {
    const xy = ChartTypes.isXY(chart.type);
    const references = plottedReferences(chart);
    if (chart.series.length === 0 || references.length === 0)
      return fail('noData', 'Add a series first: there is no data to match to a table.');
    const fallbackSheet = fallbackSheetName(workbook);
    const parsed = references.map((r) => CellRange.tryParse(r.trim()));
    const first = parsed[0];
    if (!first || parsed.some((r) => r === null))
      return fail('notRanges', 'One of the ranges isn’t a cell range.');
    const sheet = workbook.findSheet(first.sheet ?? fallbackSheet);
    const table = sheet?.tables.find((t) => t.range.contains(first.top, first.left));
    if (!sheet || !table)
      return fail('notInTable', 'The ranges aren’t inside a table. Make the data a table first.');
    const rows = dataRows(table);
    // `wholeColumns` matches each range to the table column it sits in, whatever rows it covers.
    const fits = (range: CellRange): boolean =>
      (range.sheet ?? fallbackSheet).toLowerCase() === sheet.name.toLowerCase() &&
      range.left === range.right &&
      range.left >= table.range.left &&
      range.left <= table.range.right &&
      (options.wholeColumns === true || (range.top === rows.top && range.bottom === table.range.bottom));
    const labels = tableHeaderLabels(table, sheet);
    const misfit = parsed.findIndex((r) => !fits(r!));
    if (misfit >= 0) {
      const range = parsed[misfit]!;
      const column =
        range.left === range.right && range.left >= table.range.left && range.left <= table.range.right
          ? labels[range.left - table.range.left]
          : undefined;
      const what = column
        ? `${references[misfit]!.trim()} is only part of the column “${column}”.`
        : `${references[misfit]!.trim()} isn’t one column of the table “${table.name}”.`;
      return fail(
        'notWholeColumns',
        `${what} Each range must be one whole column of the table’s data, without its header.`,
      );
    }
    const nameOf = (reference: string | null): string | null => {
      const range = reference ? CellRange.tryParse(reference.trim()) : null;
      return range ? (labels[range.left - table.range.left] ?? null) : null;
    };
    const source: ChartTableSource = {
      table: table.name,
      categoryColumn: nameOf(xy ? (chart.series[0]?.xValuesReference ?? null) : chart.categoriesReference),
      series: chart.series.map((s: ChartSeries) => ({
        column: nameOf(s.valuesReference)!,
        type: s.type,
        sizeColumn: chart.type === 'bubble' ? nameOf(s.bubbleSizeReference) : null,
      })),
    };
    const firstX = nameOf(chart.series[0]?.xValuesReference ?? null);
    const odd = xy ? chart.series.findIndex((s) => nameOf(s.xValuesReference) !== firstX) : -1;
    if (odd > 0) {
      const name = (i: number): string =>
        chart.series[i]!.name ?? nameOf(chart.series[i]!.valuesReference) ?? `Series ${i + 1}`;
      return fail(
        'mixedX',
        `“${name(odd)}” takes its X values from another column than “${name(0)}”. Every series must share one X column to plot from a table.`,
      );
    }
    return { ok: true, chart: ChartTableBinding.bind(workbook, chart, source) };
  },

  /**
   * Whether the chart panel opens in table mode: the chart is bound, or it plots nothing yet and
   * the workbook has a table to plot from.
   */
  startsInTableMode(workbook: Workbook, chart: ChartDefinition): boolean {
    if (chart.source) return true;
    return !hasData(chart) && ChartTableBinding.tableNames(workbook).length > 0;
  },

  /**
   * Range → Table for the panel's toggle: a chart with no data yet is seeded from the first table
   * (or the one named); otherwise its ranges must line up with one table (`fromRanges`).
   */
  toTable(workbook: Workbook, chart: ChartDefinition, table: string | null = null): FromRangesResult {
    if (!hasData(chart)) {
      const name = table ?? ChartTableBinding.tableNames(workbook)[0];
      const source = name ? ChartTableBinding.seed(workbook, name, chart.type) : null;
      if (!source) return fail('noTable', 'There is no table in the workbook. Make the data a table first.');
      return { ok: true, chart: ChartTableBinding.bind(workbook, chart, source) };
    }
    return ChartTableBinding.fromRanges(workbook, chart);
  },

  /** The chart plotting the table called `name` instead, seeded as a new chart would be. */
  setTable(workbook: Workbook, chart: ChartDefinition, name: string): ChartDefinition {
    const source = ChartTableBinding.seed(workbook, name, chart.type);
    return source ? ChartTableBinding.bind(workbook, chart, source) : chart;
  },

  /** The bound chart with `column` labelling the categories (or giving the X values); null for none. */
  setCategoryColumn(workbook: Workbook, chart: ChartDefinition, column: string | null): ChartDefinition {
    if (!chart.source) return chart;
    return ChartTableBinding.bind(workbook, chart, { ...cloneSource(chart.source), categoryColumn: column });
  },

  /**
   * The bound chart plotting `columns`, in the table's column order. A column already plotted keeps
   * its own settings (its "draw as" type, its bubble sizes).
   */
  setSeriesColumns(workbook: Workbook, chart: ChartDefinition, columns: readonly string[]): ChartDefinition {
    if (!chart.source) return chart;
    const source = cloneSource(chart.source);
    const order = (ChartTableBinding.columns(workbook, source.table) ?? []).map((c) => c.name.toLowerCase());
    const wanted = [...new Set(columns.map((c) => c.toLowerCase()))].sort(
      (a, b) => order.indexOf(a) - order.indexOf(b),
    );
    source.series = wanted.map((lower) => {
      const kept = source.series.find((s) => s.column.toLowerCase() === lower);
      const name = columns.find((c) => c.toLowerCase() === lower)!;
      return kept ?? { column: name, type: null, sizeColumn: null };
    });
    return ChartTableBinding.bind(workbook, chart, source);
  },

  /** The bound bubble chart with every series sized by `column`. */
  setSizeColumn(workbook: Workbook, chart: ChartDefinition, column: string | null): ChartDefinition {
    if (!chart.source) return chart;
    const source = cloneSource(chart.source);
    source.series = source.series.map((s) => ({ ...s, sizeColumn: column }));
    return ChartTableBinding.bind(workbook, chart, source);
  },

  /**
   * The chart as `type`. A bound chart keeps its columns where the new type can plot them: a pie
   * keeps its first series; a change between category and X/Y charts, or to a bubble chart without
   * sizes, picks its columns afresh from the table.
   */
  retype(workbook: Workbook, chart: ChartDefinition, type: ChartType): ChartDefinition {
    const next = cloneChart(chart);
    next.type = type;
    if (!next.source) return next;
    const familyChanged = ChartTypes.isXY(type) !== ChartTypes.isXY(chart.type);
    const needsSizes = type === 'bubble' && next.source.series.some((s) => s.sizeColumn === null);
    if (familyChanged || needsSizes) {
      const seeded = ChartTableBinding.seed(workbook, next.source.table, type);
      if (seeded) return ChartTableBinding.bind(workbook, next, seeded);
    }
    if (type === 'pie' || type === 'doughnut') next.source.series = next.source.series.slice(0, 1);
    return ChartTableBinding.bind(workbook, next, next.source);
  },

  /**
   * Where a new table would hold the chart's data, for the panel's "Make a table": the block its
   * ranges span on one sheet, plus the row of labels just above it when there is one. Null when the
   * chart is bound or plots nothing, a range isn't a single column, the ranges span sheets, or the
   * block overlaps a table.
   */
  tableRegion(
    workbook: Workbook,
    chart: ChartDefinition,
  ): { sheet: string; range: CellRange; header: string[] | null } | null {
    if (chart.source) return null;
    const references = plottedReferences(chart);
    if (references.length === 0) return null;
    const fallbackSheet = fallbackSheetName(workbook);
    const ranges: CellRange[] = [];
    for (const reference of references) {
      const range = CellRange.tryParse(reference.trim());
      if (!range || range.left !== range.right) return null;
      ranges.push(range);
    }
    const names = new Set(ranges.map((r) => (r.sheet ?? fallbackSheet).toLowerCase()));
    if (names.size !== 1) return null;
    const sheet = workbook.findSheet(ranges[0]!.sheet ?? fallbackSheet);
    if (!sheet || sheet.isChartSheet) return null;
    if (ChartTypes.isXY(chart.type)) {
      // Every series must share one X column, or the new table couldn't be plotted by name.
      const xs = new Set(chart.series.map((s) => s.xValuesReference?.trim().toLowerCase() ?? ''));
      if (xs.size > 1) return null;
    }
    let top = Math.min(...ranges.map((r) => r.top));
    const left = Math.min(...ranges.map((r) => r.left));
    const right = Math.max(...ranges.map((r) => r.right));
    // A whole column picked from its header ends at the sheet's last row: stop at the last used one.
    const used = sheet.getUsedRange();
    if (!used) return null;
    const bottom = Math.min(Math.max(...ranges.map((r) => r.bottom)), used.bottom);
    if (bottom < top) return null;
    const labelsAt = (row: number) =>
      DataRegion.firstRowIsLabels(sheet, CellRange.fromBounds(row, left, row, right));
    if (top > 0 && !labelsAt(top) && labelsAt(top - 1)) top--;
    const range = CellRange.fromBounds(top, left, bottom, right);
    if (sheet.tables.some((t) => t.range.intersects(range))) return null;
    // The labels that become the column names; null when the table gets a new header row instead.
    const header = labelsAt(top)
      ? Array.from({ length: right - left + 1 }, (_, i) => labelOf(sheet.getValue(top, left + i)).trim())
      : null;
    return { sheet: sheet.name, range, header };
  },

  /**
   * The chart types a bound chart's table can't fill, each with why, for the panel's type picker:
   * every chart needs a number column, a scatter chart two (X and Y), a bubble chart three.
   */
  unavailableTypes(workbook: Workbook, chart: ChartDefinition): Map<ChartType, string> {
    const out = new Map<ChartType, string>();
    const columns = chart.source ? ChartTableBinding.columns(workbook, chart.source.table) : null;
    if (!columns) return out;
    const numbers = columns.filter((c) => c.kind === 'number').length;
    for (const type of ChartTypes.all) {
      if (numbers === 0) out.set(type, 'The table has no column of numbers to plot.');
      else if (type === 'bubble' && numbers < 3)
        out.set(type, 'A bubble chart needs three columns of numbers: X, Y and bubble size.');
      else if (type === 'scatter' && numbers < 2)
        out.set(type, 'A scatter chart needs two columns of numbers: X and Y.');
    }
    return out;
  },

  /** The chart with its binding pointed at `to` when it was bound to `from` (ignoring case). */
  renameTable(chart: ChartDefinition, from: string, to: string): ChartDefinition {
    if (!chart.source || chart.source.table.toLowerCase() !== from.toLowerCase()) return chart;
    const next = cloneChart(chart);
    next.source!.table = to;
    return next;
  },
} as const;
