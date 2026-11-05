import type { ChartDefinition, ChartSeries } from '@cellestra/charts';
import { createChart, createSeries } from '@cellestra/charts';
import type { Table } from '../data/types.ts';
import { quoteSheetIfNeeded } from '../model/cellAddress.ts';
import type { CellRange } from '../model/cellRange.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { ActionCommand, type EditCommand, EditResult, requireSheet } from './editCommand.ts';

const letters = (column: number): string => SheetDimensions.columnToLetters(column);

/** `Sheet.B2:B7`: one column's rows `fromRow..toRow` (zero-based) on `sheet`. */
function columnRef(sheet: string, column: number, fromRow: number, toRow: number): string {
  return `${quoteSheetIfNeeded(sheet)}.${letters(column)}${fromRow + 1}:${letters(column)}${toRow + 1}`;
}

function columnHasNumber(sheet: Worksheet, column: number, fromRow: number, toRow: number): boolean {
  for (let row = fromRow; row <= toRow; row++) {
    if (sheet.getValue(row, column).kind === 'number') return true;
  }
  return false;
}

/**
 * Seed from the sheet's single table: its first column labels the categories and its first numeric
 * column becomes one series (named from the header row when present). Null when the table can't be
 * charted (header-only, or no numeric column beside the labels).
 */
function seriesFromTable(
  sheet: Worksheet,
  table: Table,
): { categories: string; series: ChartSeries[] } | null {
  const range = table.range;
  const dataTop = table.hasHeader ? range.top + 1 : range.top;
  if (dataTop > range.bottom) return null;
  let valueColumn: number | null = null;
  for (let column = range.left + 1; column <= range.right; column++) {
    if (columnHasNumber(sheet, column, dataTop, range.bottom)) {
      valueColumn = column;
      break;
    }
  }
  if (valueColumn === null) return null;
  const series = createSeries({ valuesReference: columnRef(sheet.name, valueColumn, dataTop, range.bottom) });
  if (table.hasHeader)
    series.nameReference = `${quoteSheetIfNeeded(sheet.name)}.${letters(valueColumn)}${range.top + 1}`;
  return { categories: columnRef(sheet.name, range.left, dataTop, range.bottom), series: [series] };
}

/**
 * Seed from a selection that looks tabular: the first column labels the categories and each other
 * column becomes a series, named from a text header row when present. A single column becomes one
 * unnamed series with no categories. Nothing when fewer than two rows are selected.
 */
function seedFromSelection(definition: ChartDefinition, sheet: Worksheet, range: CellRange): void {
  if (range.rowCount < 2) return;
  const { top, bottom, left, right } = range;
  // A header row has a text label over the value columns (the top-left corner is often blank).
  let headerRow = false;
  for (let column = right > left ? left + 1 : left; column <= right && !headerRow; column++)
    headerRow = sheet.getValue(top, column).kind === 'text';
  const dataTop = headerRow && bottom > top ? top + 1 : top;
  if (right > left) {
    definition.categoriesReference = columnRef(sheet.name, left, dataTop, bottom);
    for (let column = left + 1; column <= right; column++) {
      const series = createSeries({ valuesReference: columnRef(sheet.name, column, dataTop, bottom) });
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
   * selection). When the sheet has exactly one chartable table, the table wins; otherwise the
   * selection seeds it. A chart sheet or a missing sheet gives an empty column chart.
   */
  chartFromRange(workbook: Workbook, sheetName: string, range: CellRange): ChartDefinition {
    const definition = createChart({ type: 'column', title: 'Chart', series: [] });
    const sheet = workbook.findSheet(sheetName);
    if (!sheet || sheet.isChartSheet) return definition;
    const onlyTable = sheet.tables.length === 1 ? sheet.tables[0] : undefined;
    const fromTable = onlyTable ? seriesFromTable(sheet, onlyTable) : null;
    if (fromTable) {
      definition.categoriesReference = fromTable.categories;
      definition.series.push(...fromTable.series);
    } else {
      seedFromSelection(definition, sheet, range);
    }
    return definition;
  },
} as const;
