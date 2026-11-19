import type { ChartDefinition } from '@cellestra/charts';
import { ChartTypes, createChart, createSeries } from '@cellestra/charts';
import { ChartTableBinding } from '../charts/chartTableBinding.ts';
import { quoteSheetIfNeeded } from '../model/cellAddress.ts';
import type { CellRange } from '../model/cellRange.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { ActionCommand, type EditCommand, EditResult, requireSheet } from './editCommand.ts';
import { clampToUsed, EditCommands, OVERSIZED_CELL_CAP } from './editCommands.ts';
import { tableNameError, TableCommands } from './tableCommands.ts';

const letters = (column: number): string => SheetDimensions.columnToLetters(column);

/** `Sheet.B2:B7`: one column's rows `fromRow..toRow` (zero-based) on `sheet`. */
function columnRef(sheet: string, column: number, fromRow: number, toRow: number): string {
  return `${quoteSheetIfNeeded(sheet)}.${letters(column)}${fromRow + 1}:${letters(column)}${toRow + 1}`;
}

/**
 * Seed `definition` from a range that looks tabular: the first column labels the categories (the X
 * values, for a scatter or bubble chart) and each other column becomes a series, named from a text
 * header row when present. A single column becomes one unnamed series with no categories. Nothing
 * when fewer than two rows are given. A huge range (whole columns, the whole sheet) is cut to the
 * used cells first, or it would chart every row. Insert ▸ Chart Sheet and the CLI's `chart --from`
 * share it.
 */
export function seedChartFromRange(definition: ChartDefinition, sheet: Worksheet, selected: CellRange): void {
  const range = selected.cellCount < OVERSIZED_CELL_CAP ? selected : clampToUsed(sheet, selected);
  if (!range || range.rowCount < 2) return;
  const xy = ChartTypes.isXY(definition.type);
  const { top, bottom, left, right } = range;
  // A header row has a text label over the value columns (the top-left corner is often blank).
  let headerRow = false;
  for (let column = right > left ? left + 1 : left; column <= right && !headerRow; column++)
    headerRow = sheet.getValue(top, column).kind === 'text';
  const dataTop = headerRow && bottom > top ? top + 1 : top;
  if (right > left) {
    const firstColumn = columnRef(sheet.name, left, dataTop, bottom);
    if (!xy) definition.categoriesReference = firstColumn;
    for (let column = left + 1; column <= right; column++) {
      const series = createSeries({
        valuesReference: columnRef(sheet.name, column, dataTop, bottom),
        xValuesReference: xy ? firstColumn : null,
      });
      if (headerRow) series.nameReference = `${quoteSheetIfNeeded(sheet.name)}.${letters(column)}${top + 1}`;
      definition.series.push(series);
    }
  } else {
    definition.series.push(createSeries({ valuesReference: columnRef(sheet.name, left, dataTop, bottom) }));
  }
}

/** The edit that makes a chart's data a table and binds the chart to it, or why it can't. */
export type ChartTablePlan =
  { readonly ok: true; readonly command: EditCommand } | { readonly ok: false; readonly reason: string };

export const ChartEditCommands = {
  /** Replace a chart sheet's definition (undo restores the previous one). */
  setChart(sheet: string, next: ChartDefinition): EditCommand {
    let previous: ChartDefinition | null = null;
    return new ActionCommand(
      'Edit Chart',
      (wb) => {
        const target = requireSheet(wb, sheet);
        previous = target.chart;
        target.chart = next;
        return EditResult.none;
      },
      (wb) => {
        requireSheet(wb, sheet).chart = previous;
        return EditResult.none;
      },
    );
  },

  /**
   * "Make a table" from the chart panel, as one undo step: the block the chart's ranges cover
   * (`ChartTableBinding.tableRegion`) becomes a table called `name`, with a header row of labels
   * (one is added when the data has none), and the chart on `chartSheet` then plots the table's
   * columns by name. Each range binds to the column it sits in, so a range a row short still binds.
   */
  makeTableForChart(workbook: Workbook, chartSheet: string, name: string): ChartTablePlan {
    const chart = workbook.findSheet(chartSheet)?.chart ?? null;
    const region = chart ? ChartTableBinding.tableRegion(workbook, chart) : null;
    if (!region) return { ok: false, reason: 'The chart’s data can’t be made into one table.' };
    const error = tableNameError(workbook, name.trim());
    if (error) return { ok: false, reason: error };
    const plan = TableCommands.create(workbook, region.sheet, region.range, {
      header: 'labels',
      name: name.trim(),
    });
    if (!plan.ok) return { ok: false, reason: 'The chart’s data overlaps a table.' };
    let previous: ChartDefinition | null = null;
    const bind = new ActionCommand(
      'Plot Table Columns',
      (wb) => {
        // The table exists now, and a header row it inserted has moved the chart's ranges down.
        const target = requireSheet(wb, chartSheet);
        previous = target.chart;
        const result = previous ? ChartTableBinding.fromRanges(wb, previous, { wholeColumns: true }) : null;
        if (result?.ok) target.chart = result.chart;
        return EditResult.none;
      },
      (wb) => {
        requireSheet(wb, chartSheet).chart = previous;
        return EditResult.none;
      },
    );
    return { ok: true, command: EditCommands.composite('Make Table', [plan.command, bind]) };
  },

  /**
   * The default chart for a new chart sheet, built from `range` on `sheetName` (the editor's
   * selection). When the selection is in a table (or the sheet has just one table), the chart is
   * bound to that table: its first text column labels the categories and its number columns become
   * series (`ChartTableBinding.seed`). Otherwise, or when the table has no number column, the
   * selection seeds plain ranges. A chart sheet or a missing sheet gives an empty column chart.
   */
  chartFromRange(workbook: Workbook, sheetName: string, range: CellRange): ChartDefinition {
    const definition = createChart({ type: 'column', title: 'Chart', series: [] });
    const sheet = workbook.findSheet(sheetName);
    if (!sheet || sheet.isChartSheet) return definition;
    const table =
      sheet.tables.find((t) => t.range.contains(range.top, range.left)) ??
      (sheet.tables.length === 1 ? sheet.tables[0] : undefined);
    const source = table ? ChartTableBinding.seed(workbook, table.name, definition.type) : null;
    if (source && source.series.length > 0) return ChartTableBinding.bind(workbook, definition, source);
    seedChartFromRange(definition, sheet, range);
    return definition;
  },
} as const;
