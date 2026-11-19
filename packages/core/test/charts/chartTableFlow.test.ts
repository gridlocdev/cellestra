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
    });
    expect(ChartTableBinding.toTable(Workbook.createNew(), createChart())).toMatchObject({ code: 'noTable' });
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
    });
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
