import type { ChartDefinition } from '@cellestra/charts';
import { ChartTypes, createChart, createSeries } from '@cellestra/charts';
import { ChartTableBinding } from '../charts/chartTableBinding.ts';
import { quoteSheetIfNeeded } from '../model/cellAddress.ts';
import type { CellRange } from '../model/cellRange.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { ActionCommand, type EditCommand, EditResult, requireSheet } from './editCommand.ts';
import { clampToUsed, OVERSIZED_CELL_CAP } from './editCommands.ts';

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
