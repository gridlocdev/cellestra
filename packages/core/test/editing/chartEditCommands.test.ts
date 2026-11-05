import { createChart, createSeries } from '@cellestra/charts';
import { describe, expect, it } from 'vitest';
import { ChartEditCommands } from '../../src/editing/chartEditCommands.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const R = (a1: string): CellRange => CellRange.parse(a1);

/** Sheet1: A1:C4 = Month | Sales | Cost, Jan..Mar. */
function data(): Workbook {
  const wb = Workbook.createNew();
  const s = wb.activeSheet;
  ['Month', 'Sales', 'Cost'].forEach((h, c) => s.setValue(0, c, CellValue.text(h)));
  ['Jan', 'Feb', 'Mar'].forEach((m, r) => {
    s.setValue(r + 1, 0, CellValue.text(m));
    s.setValue(r + 1, 1, CellValue.number(10 * (r + 1)));
    s.setValue(r + 1, 2, CellValue.number(r + 1));
  });
  return wb;
}

describe('ChartEditCommands.setChart', () => {
  it('swaps the chart in and undo restores the previous one', () => {
    const wb = data();
    const before = createChart({ title: 'Before' });
    const sheet = wb.addChartSheet(before);
    const after = createChart({ title: 'After', type: 'line' });
    const cmd = ChartEditCommands.setChart(sheet.name, after);

    const result = cmd.apply(wb);
    expect(result.structural).toBe(false);
    expect(result.changedCells).toEqual([]);
    expect(sheet.chart).toBe(after);

    cmd.undo(wb);
    expect(sheet.chart).toBe(before);
    cmd.apply(wb); // redo
    expect(sheet.chart).toBe(after);
  });

  it('works through the edit history', () => {
    const wb = data();
    const sheet = wb.addChartSheet(createChart({ title: 'A' }));
    const history = new EditHistory(wb);
    history.execute(ChartEditCommands.setChart(sheet.name, createChart({ title: 'B' })));
    expect(sheet.chart?.title).toBe('B');
    history.undo();
    expect(sheet.chart?.title).toBe('A');
  });
});

describe('ChartEditCommands.chartFromRange', () => {
  it('a header row names each value column; the first column labels the categories', () => {
    const chart = ChartEditCommands.chartFromRange(data(), 'Sheet1', R('A1:C4'));
    expect(chart.type).toBe('column');
    expect(chart.title).toBe('Chart');
    expect(chart.categoriesReference).toBe('Sheet1.A2:A4');
    expect(chart.series).toEqual([
      createSeries({ valuesReference: 'Sheet1.B2:B4', nameReference: 'Sheet1.B1' }),
      createSeries({ valuesReference: 'Sheet1.C2:C4', nameReference: 'Sheet1.C1' }),
    ]);
  });

  it('without a text header every selected row is data', () => {
    const chart = ChartEditCommands.chartFromRange(data(), 'Sheet1', R('B2:C4'));
    expect(chart.categoriesReference).toBe('Sheet1.B2:B4');
    expect(chart.series).toEqual([createSeries({ valuesReference: 'Sheet1.C2:C4' })]);
  });

  it('a single column becomes one unnamed series with no categories', () => {
    const chart = ChartEditCommands.chartFromRange(data(), 'Sheet1', R('B2:B4'));
    expect(chart.categoriesReference).toBeNull();
    expect(chart.series).toEqual([createSeries({ valuesReference: 'Sheet1.B2:B4' })]);
  });

  it('fewer than two rows gives an empty chart', () => {
    const chart = ChartEditCommands.chartFromRange(data(), 'Sheet1', R('A1:C1'));
    expect(chart.series).toEqual([]);
    expect(chart.categoriesReference).toBeNull();
  });

  it("the sheet's single table wins over the selection", () => {
    const wb = data();
    wb.activeSheet.addTable({ name: 'T', range: R('A1:C4'), hasHeader: true });
    const chart = ChartEditCommands.chartFromRange(wb, 'Sheet1', R('D9'));
    expect(chart.categoriesReference).toBe('Sheet1.A2:A4');
    expect(chart.series).toEqual([
      createSeries({ valuesReference: 'Sheet1.B2:B4', nameReference: 'Sheet1.B1' }),
    ]);
  });

  it('a table with no numeric column falls back to the selection', () => {
    const wb = data();
    wb.activeSheet.addTable({ name: 'T', range: R('A1:A4'), hasHeader: true });
    const chart = ChartEditCommands.chartFromRange(wb, 'Sheet1', R('B2:B4'));
    expect(chart.categoriesReference).toBeNull();
    expect(chart.series).toEqual([createSeries({ valuesReference: 'Sheet1.B2:B4' })]);
  });

  it('a chart sheet or missing sheet gives an empty chart', () => {
    const wb = data();
    const chartSheet = wb.addChartSheet(createChart());
    expect(ChartEditCommands.chartFromRange(wb, chartSheet.name, R('A1:C4')).series).toEqual([]);
    expect(ChartEditCommands.chartFromRange(wb, 'Nope', R('A1:C4')).series).toEqual([]);
  });
});
