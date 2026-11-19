import type { ChartDefinition } from '@cellestra/charts';
import { ChartTypes } from '@cellestra/charts';
import { quoteSheetIfNeeded } from '../model/cellAddress.ts';
import { findTableByName } from '../data/tables.ts';
import { CellRange } from '../model/cellRange.ts';
import type { Workbook } from '../model/workbook.ts';
import { labelOf } from './chartLabels.ts';
import type { ChartField } from './chartFields.ts';
import { fallbackSheetName, plottedReferences } from './chartTableBinding.ts';

/** A reference field of a chart that can't plot as typed, and what to do about it. */
export interface ChartFieldProblem {
  readonly field: ChartField;
  readonly message: string;
}

/** Above this many cells a length check would only compare whole-column references. */
const LENGTH_CHECK_LIMIT = 100_000;

type Shape = 'cell' | 'line';

/** `A, B and C`. */
function listOf(names: readonly string[]): string {
  if (names.length <= 1) return names[0] ?? '';
  return `${names.slice(0, -1).join(', ')} and ${names[names.length - 1]}`;
}

/** How many cells a sample shows before it says how many there are. */
const SAMPLE_CELLS = 3;

/**
 * What the chart panel says about a chart's data: the problems in the ranges typed into it, so a
 * field explains why the chart is empty instead of plotting nothing, and a one-line summary of
 * what the chart plots.
 */
export const ChartChecks = {
  /**
   * What `reference` points at, for the line under its field: its first few cells (`January,
   * February, March … 12 cells`), or "Empty cells". Null when it is blank or can't be read, which
   * `problems` explains instead.
   */
  sample(workbook: Workbook, reference: string | null): string | null {
    if (!reference || reference.trim() === '') return null;
    const range = CellRange.tryParse(reference.trim());
    const sheet = range ? workbook.findSheet(range.sheet ?? fallbackSheetName(workbook)) : null;
    if (!range || !sheet || sheet.isChartSheet) return null;
    if (range.rowCount > 1 && range.columnCount > 1) return null;
    const cells = range.cellCount;
    const shown: string[] = [];
    for (let i = 0; i < Math.min(cells, SAMPLE_CELLS); i++) {
      const row = range.rowCount > 1 ? range.top + i : range.top;
      const column = range.rowCount > 1 ? range.left : range.left + i;
      shown.push(labelOf(sheet.getValue(row, column)).trim() || '(empty)');
    }
    if (shown.every((t) => t === '(empty)') && cells <= SAMPLE_CELLS)
      return cells === 1 ? 'An empty cell' : 'Empty cells';
    return cells > SAMPLE_CELLS ? `${shown.join(', ')} … ${cells} cells` : shown.join(', ');
  },

  /**
   * Every reference field that can't plot as typed: not a cell range, a sheet that doesn't exist
   * or holds a chart, a block where one row or column is needed, a name cell that is more than one
   * cell, and values that don't line up one-for-one with the labels or X values. Blank fields are
   * fine. A table-bound chart has none, because its ranges come from the table.
   */
  problems(workbook: Workbook, chart: ChartDefinition): ChartFieldProblem[] {
    if (chart.source) return [];
    const fallback = fallbackSheetName(workbook);
    const out: ChartFieldProblem[] = [];
    /** The parsed range in `text`, after reporting what's wrong with it; null when blank or broken. */
    const check = (field: ChartField, text: string | null, shape: Shape): CellRange | null => {
      if (!text || text.trim() === '') return null;
      const problem = (message: string) => {
        out.push({ field, message });
        return null;
      };
      const range = CellRange.tryParse(text.trim());
      if (!range) return problem(`“${text.trim()}” isn’t a cell range. Write it like Sheet1.B2:B7.`);
      const name = range.sheet ?? fallback;
      const sheet = workbook.findSheet(name);
      if (!sheet) return problem(`There is no sheet called “${name}”.`);
      if (sheet.isChartSheet)
        return problem(`“${sheet.name}” is a chart sheet. Use cells on a sheet with data.`);
      if (shape === 'cell' && !range.isSingleCell) return problem('A name cell is one cell, like Sheet1.B1.');
      if (shape === 'line' && range.rowCount > 1 && range.columnCount > 1)
        return problem('Use one column or one row of cells, not a block.');
      return range;
    };
    const count = (range: CellRange | null) =>
      range && range.cellCount <= LENGTH_CHECK_LIMIT ? range.cellCount : null;
    const xy = ChartTypes.isXY(chart.type);
    const categories = xy ? null : check({ kind: 'categories' }, chart.categoriesReference, 'line');
    const labels = count(categories);
    chart.series.forEach((s, i) => {
      check({ kind: 'nameCell', series: i }, s.nameReference, 'cell');
      const xValues = xy ? check({ kind: 'xValues', series: i }, s.xValuesReference, 'line') : null;
      const values = count(check({ kind: 'values', series: i }, s.valuesReference, 'line'));
      const sizes =
        chart.type === 'bubble'
          ? count(check({ kind: 'bubbleSizes', series: i }, s.bubbleSizeReference, 'line'))
          : null;
      const xs = count(xValues);
      if (values !== null && labels !== null && values !== labels)
        out.push({
          field: { kind: 'values', series: i },
          message: `These ${values} values don’t match the ${labels} category labels one for one.`,
        });
      if (values !== null && xs !== null && values !== xs)
        out.push({
          field: { kind: 'values', series: i },
          message: `These ${values} Y values don’t match the ${xs} X values one for one.`,
        });
      if (sizes !== null && values !== null && sizes !== values)
        out.push({
          field: { kind: 'bubbleSizes', series: i },
          message: `These ${sizes} sizes don’t match the ${values} Y values one for one.`,
        });
    });
    return out;
  },

  /**
   * One line naming where the chart's data comes from: `From table “Sales”: Sales and Cost by
   * Region`, or `From cells Data.A2:C7`.
   */
  summary(workbook: Workbook, chart: ChartDefinition): string {
    const source = chart.source;
    if (source) {
      if (!findTableByName(workbook, source.table))
        return `From table “${source.table}”, which no longer exists.`;
      const plotted = listOf(source.series.map((s) => s.column));
      const head = `From table “${source.table}”`;
      if (plotted === '') return `${head}: no columns picked yet.`;
      if (source.categoryColumn === null) return `${head}: ${plotted}, one point per row.`;
      return ChartTypes.isXY(chart.type)
        ? `${head}: ${plotted} against ${source.categoryColumn}.`
        : `${head}: ${plotted} by ${source.categoryColumn}.`;
    }
    const ranges = plottedReferences(chart)
      .map((r) => CellRange.tryParse(r.trim()))
      .filter((r): r is CellRange => r !== null);
    if (ranges.length === 0) return 'No data yet.';
    const fallback = fallbackSheetName(workbook);
    const sheets = new Set(ranges.map((r) => r.sheet ?? fallback));
    if (sheets.size > 1) return `From cells on ${sheets.size} sheets.`;
    const box = CellRange.fromBounds(
      Math.min(...ranges.map((r) => r.top)),
      Math.min(...ranges.map((r) => r.left)),
      Math.max(...ranges.map((r) => r.bottom)),
      Math.max(...ranges.map((r) => r.right)),
    );
    return `From cells ${quoteSheetIfNeeded([...sheets][0]!)}.${box.toString()}.`;
  },
} as const;
