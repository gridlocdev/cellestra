import { createChart, createSeries } from '@cellestra/charts';
import { describe, expect, it } from 'vitest';
import {
  CellRange,
  CellValue,
  ChartEditCommands,
  ChartTableBinding,
  EditHistory,
  table as makeTable,
  Workbook,
} from '../../src/index.ts';

/** Data.A1:C4 = Region | Sales | Cost over North/South/East, not a table. */
function book(header = true): Workbook {
  const wb = Workbook.createNew();
  const s = wb.activeSheet;
  s.name = 'Data';
  const top = header ? 1 : 0;
  if (header) ['Region', 'Sales', 'Cost'].forEach((h, c) => s.setValue(0, c, CellValue.text(h)));
  [
    ['North', 10, 1],
    ['South', 20, 2],
    ['East', 30, 3],
  ].forEach((row, r) =>
    row.forEach((v, c) =>
      s.setValue(r + top, c, typeof v === 'number' ? CellValue.number(v) : CellValue.text(v)),
    ),
  );
  return wb;
}

/** A column chart of Data.B and C by A, rows `from..to` (one-based). */
function rangeChart(from = 2, to = 4) {
  return createChart({
    categoriesReference: `Data.A${from}:A${to}`,
    series: [
      createSeries({ valuesReference: `Data.B${from}:B${to}`, nameReference: 'Data.B1' }),
      createSeries({ valuesReference: `Data.C${from}:C${to}`, nameReference: 'Data.C1', type: 'line' }),
    ],
  });
}

describe('ChartTableBinding.fromRanges failure codes', () => {
  it('names why ranges can’t bind, so the panel can offer the fix', () => {
    const wb = book();
    expect(ChartTableBinding.fromRanges(wb, createChart())).toMatchObject({ ok: false, code: 'noData' });
    expect(ChartTableBinding.fromRanges(wb, rangeChart())).toMatchObject({ ok: false, code: 'notInTable' });
    wb.findSheet('Data')!.addTable(makeTable('Sales', CellRange.parse('A1:C4')));
    expect(ChartTableBinding.fromRanges(wb, rangeChart(2, 3))).toMatchObject({
      ok: false,
      code: 'notWholeColumns',
      reason: expect.stringContaining('Data.A2:A3 is only part of the column “Region”.'),
    });
    expect(ChartTableBinding.toTable(Workbook.createNew(), createChart())).toMatchObject({ code: 'noTable' });
  });

  it('names the series whose X column differs', () => {
    const wb = book();
    wb.findSheet('Data')!.addTable(makeTable('Sales', CellRange.parse('A1:C4')));
    const chart = createChart({
      type: 'scatter',
      series: [
        createSeries({ name: 'Sales', xValuesReference: 'Data.A2:A4', valuesReference: 'Data.B2:B4' }),
        createSeries({ name: 'Cost', xValuesReference: 'Data.B2:B4', valuesReference: 'Data.C2:C4' }),
      ],
    });
    expect(ChartTableBinding.fromRanges(wb, chart)).toMatchObject({
      code: 'mixedX',
      reason: expect.stringMatching(/^“Cost” takes its X values from another column than “Sales”\./),
    });
  });

  it('with wholeColumns, binds each range to the column it sits in', () => {
    const wb = book();
    wb.findSheet('Data')!.addTable(makeTable('Sales', CellRange.parse('A1:C4')));
    const result = ChartTableBinding.fromRanges(wb, rangeChart(2, 3), { wholeColumns: true });
    if (!result.ok) throw new Error(result.reason);
    expect(result.chart.source?.series.map((s) => s.column)).toEqual(['Sales', 'Cost']);
    expect(result.chart.series[0]!.valuesReference).toBe('Data.B2:B4');
  });
});

describe('ChartTableBinding.tableRegion', () => {
  it('spans the chart’s ranges plus the label row above them', () => {
    expect(ChartTableBinding.tableRegion(book(), rangeChart())).toEqual({
      sheet: 'Data',
      range: CellRange.fromBounds(0, 0, 3, 2),
      header: ['Region', 'Sales', 'Cost'],
    });
    expect(ChartTableBinding.tableRegion(book(false), rangeChart(1, 3))?.header).toBeNull();
  });

  it('keeps a block with no labels above as it is', () => {
    expect(ChartTableBinding.tableRegion(book(false), rangeChart(1, 3))?.range).toEqual(
      CellRange.fromBounds(0, 0, 2, 2),
    );
  });

  it('is null for bound, empty, block, cross-sheet or table-overlapping charts', () => {
    const wb = book();
    expect(ChartTableBinding.tableRegion(wb, createChart())).toBeNull();
    const block = createChart({ series: [createSeries({ valuesReference: 'Data.B2:C4' })] });
    expect(ChartTableBinding.tableRegion(wb, block)).toBeNull();
    wb.addSheet('Other');
    const split = createChart({
      categoriesReference: 'Other.A2:A4',
      series: [createSeries({ valuesReference: 'Data.B2:B4' })],
    });
    expect(ChartTableBinding.tableRegion(wb, split)).toBeNull();
    wb.findSheet('Data')!.addTable(makeTable('Sales', CellRange.parse('C1:C4')));
    expect(ChartTableBinding.tableRegion(wb, rangeChart())).toBeNull();
  });
});

describe('ChartTableBinding.unavailableTypes', () => {
  it('rules out the types the table’s number columns can’t fill', () => {
    const wb = book();
    wb.findSheet('Data')!.addTable(makeTable('Sales', CellRange.parse('A1:C4')));
    const chart = ChartTableBinding.setTable(wb, createChart(), 'Sales');
    expect([...ChartTableBinding.unavailableTypes(wb, chart).keys()]).toEqual(['bubble']);
    expect(ChartTableBinding.unavailableTypes(wb, rangeChart()).size).toBe(0);
  });
});

describe('ChartEditCommands.makeTableForChart', () => {
  it('makes the data a table and binds the chart to it, in one undo step', () => {
    const wb = book();
    const sheet = wb.addChartSheet(rangeChart());
    const plan = ChartEditCommands.makeTableForChart(wb, sheet.name, 'Sales');
    if (!plan.ok) throw new Error(plan.reason);
    const history = new EditHistory(wb);
    history.execute(plan.command);
    expect(wb.findSheet('Data')!.tables.map((t) => [t.name, t.range.toString()])).toEqual([
      ['Sales', 'A1:C4'],
    ]);
    expect(sheet.chart?.source).toEqual({
      table: 'Sales',
      categoryColumn: 'Region',
      series: [
        { column: 'Sales', type: null, sizeColumn: null },
        { column: 'Cost', type: 'line', sizeColumn: null },
      ],
    });
    history.undo();
    expect(wb.findSheet('Data')!.tables).toEqual([]);
    expect(sheet.chart?.source).toBeNull();
    expect(sheet.chart?.series[0]!.valuesReference).toBe('Data.B2:B4');
  });

  it('adds a header row to data without one, and the chart still binds', () => {
    const wb = book(false);
    const sheet = wb.addChartSheet(rangeChart(1, 3));
    const plan = ChartEditCommands.makeTableForChart(wb, sheet.name, 'Sales');
    if (!plan.ok) throw new Error(plan.reason);
    new EditHistory(wb).execute(plan.command);
    expect(sheet.chart?.source?.series).toHaveLength(2);
    expect(sheet.chart?.series[0]!.valuesReference).toBe('Data.B2:B4');
  });

  it('refuses a bad name or a chart whose data can’t be one table', () => {
    const wb = book();
    const sheet = wb.addChartSheet(rangeChart());
    expect(ChartEditCommands.makeTableForChart(wb, sheet.name, 'A1')).toMatchObject({ ok: false });
    const empty = wb.addChartSheet(createChart());
    expect(ChartEditCommands.makeTableForChart(wb, empty.name, 'Sales')).toMatchObject({ ok: false });
  });
});

describe('Edit Data Range', () => {
  it('re-plots a range chart from one block, keeping its type and look', () => {
    const wb = book();
    const chart = { ...rangeChart(), type: 'line' as const, title: 'Costs', showLegend: false };
    const next = ChartEditCommands.withDataRange(wb, chart, 'Data', CellRange.parse('A1:B4'));
    expect(next).toMatchObject({ type: 'line', title: 'Costs', showLegend: false, source: null });
    expect(next!.categoriesReference).toBe('Data.A2:A4');
    expect(next!.series.map((s) => s.valuesReference)).toEqual(['Data.B2:B4']);
    expect(next!.series[0]!.nameReference).toBe('Data.B1');
  });

  it('plots the block picked in a table, binding by name only when it takes whole columns', () => {
    const wb = book();
    wb.findSheet('Data')!.addTable(makeTable('Sales', CellRange.parse('A1:C4')));
    const part = ChartEditCommands.withDataRange(wb, rangeChart(), 'Data', CellRange.parse('A1:B3'));
    expect(part).toMatchObject({ source: null, categoriesReference: 'Data.A2:A3' });
    expect(part!.series.map((s) => s.valuesReference)).toEqual(['Data.B2:B3']);
    const whole = ChartEditCommands.withDataRange(wb, rangeChart(), 'Data', CellRange.parse('A1:B4'));
    expect(whole?.source).toMatchObject({ table: 'Sales', categoryColumn: 'Region' });
    expect(whole!.source!.series.map((s) => s.column)).toEqual(['Sales']);
    expect(ChartEditCommands.withDataRange(wb, rangeChart(), 'Data', CellRange.parse('F1:F1'))).toBeNull();
  });

  it('starts the pick on the chart’s block, its table, or its sheet', () => {
    const wb = book();
    expect(ChartEditCommands.dataRangeTarget(wb, rangeChart())).toEqual({
      sheet: 'Data',
      range: CellRange.fromBounds(0, 0, 3, 2),
    });
    expect(ChartEditCommands.dataRangeTarget(wb, createChart())).toEqual({ sheet: 'Data', range: null });
  });
});

describe('Make a Table region', () => {
  it('stops a column picked from its header at the last used row', () => {
    const wb = book();
    const chart = createChart({
      categoriesReference: 'Data.A2:A1048576',
      series: [createSeries({ valuesReference: 'Data.B2:B1048576' })],
    });
    expect(ChartTableBinding.tableRegion(wb, chart)?.range).toEqual(CellRange.fromBounds(0, 0, 3, 1));
  });

  it('isn’t offered for an X/Y chart whose series take X from different columns', () => {
    const wb = book();
    const chart = createChart({
      type: 'scatter',
      series: [
        createSeries({ valuesReference: 'Data.B2:B4', xValuesReference: 'Data.A2:A4' }),
        createSeries({ valuesReference: 'Data.C2:C4', xValuesReference: 'Data.B2:B4' }),
      ],
    });
    expect(ChartTableBinding.tableRegion(wb, chart)).toBeNull();
  });
});
