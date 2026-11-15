import { createChart, createSeries, type ChartTableSource } from '@cellestra/charts';
import { describe, expect, it } from 'vitest';
import {
  CellRange,
  CellValue,
  ChartDataResolver,
  ChartTableBinding,
  EditHistory,
  readOdsWithPackage,
  StructuralEditCommand,
  table as makeTable,
  TableCommands,
  Workbook,
  writeOds,
} from '../../src/index.ts';

const R = (a1: string): CellRange => CellRange.parse(a1);

/** Data.A1:C4 = Region | Sales | Cost, North/South/East, as the table `Sales`. */
function book(): Workbook {
  const wb = Workbook.createNew();
  const s = wb.activeSheet;
  s.name = 'Data';
  ['Region', 'Sales', 'Cost'].forEach((h, c) => s.setValue(0, c, CellValue.text(h)));
  [
    ['North', 10, 1],
    ['South', 20, 2],
    ['East', 30, 3],
  ].forEach((row, r) =>
    row.forEach((v, c) =>
      s.setValue(r + 1, c, typeof v === 'number' ? CellValue.number(v) : CellValue.text(v)),
    ),
  );
  s.addTable(makeTable('Sales', R('A1:C4')));
  return wb;
}

const source = (series: string[], categoryColumn: string | null = 'Region'): ChartTableSource => ({
  table: 'Sales',
  categoryColumn,
  series: series.map((column) => ({ column, type: null, sizeColumn: null })),
});

function bound(wb: Workbook, series = ['Sales']) {
  return ChartTableBinding.bind(wb, createChart({ type: 'column' }), source(series));
}

describe('ChartTableBinding.materialize', () => {
  it('turns column names into the table’s current ranges', () => {
    const chart = bound(book(), ['Sales', 'Cost']);
    expect(chart.categoriesReference).toBe('Data.A2:A4');
    expect(chart.series.map((s) => [s.valuesReference, s.nameReference])).toEqual([
      ['Data.B2:B4', 'Data.B1'],
      ['Data.C2:C4', 'Data.C1'],
    ]);
  });

  it('follows the table as it grows', () => {
    const wb = book();
    const chart = bound(wb);
    const sheet = wb.activeSheet;
    sheet.setValue(4, 0, CellValue.text('West'));
    sheet.setValue(4, 1, CellValue.number(40));
    new EditHistory(wb).execute(TableCommands.expandTableForEdit(wb, 'Data', sheet.tables[0]!, 4, 0)!);
    const resolved = ChartDataResolver.resolve(wb, chart);
    expect(resolved.categories).toEqual(['North', 'South', 'East', 'West']);
    expect(resolved.series[0]!.values).toEqual([10, 20, 30, 40]);
  });

  it('finds a column by name after a column is inserted before it', () => {
    const wb = book();
    const chart = bound(wb, ['Cost']);
    new EditHistory(wb).execute(new StructuralEditCommand('Data', 'column', 'insert', 1, 1));
    wb.activeSheet.setValue(0, 1, CellValue.text('Notes'));
    expect(ChartDataResolver.resolve(wb, chart).series[0]!.values).toEqual([1, 2, 3]);
    expect(ChartTableBinding.materialize(wb, chart).series[0]!.valuesReference).toBe('Data.D2:D4');
  });

  it('keeps a renamed column by where it was, and takes its new name', () => {
    const wb = book();
    const chart = bound(wb, ['Cost']);
    wb.activeSheet.setValue(0, 2, CellValue.text('Spend'));
    const next = ChartTableBinding.materialize(wb, chart);
    expect(next.series[0]!.valuesReference).toBe('Data.C2:C4');
    expect(next.source!.series[0]!.column).toBe('Spend');
    expect(ChartDataResolver.resolve(wb, chart).series[0]!.name).toBe('Spend');
  });

  it('falls back to the stored ranges when the table is gone', () => {
    const wb = book();
    const chart = bound(wb);
    wb.activeSheet.removeTable(wb.activeSheet.tables[0]!);
    wb.activeSheet.setValue(4, 1, CellValue.number(99)); // outside the stored range
    expect(ChartTableBinding.materialize(wb, chart)).toBe(chart);
    expect(ChartDataResolver.resolve(wb, chart).series[0]!.values).toEqual([10, 20, 30]);
  });

  it('plots X/Y charts from the category column as X', () => {
    const wb = book();
    const chart = ChartTableBinding.bind(wb, createChart({ type: 'scatter' }), source(['Cost'], 'Sales'));
    expect(chart.categoriesReference).toBeNull();
    expect(chart.series[0]!.xValuesReference).toBe('Data.B2:B4');
    expect(chart.series[0]!.valuesReference).toBe('Data.C2:C4');
  });
});

describe('Renaming a bound table', () => {
  it('points the chart at the new name, and undo points it back', () => {
    const wb = book();
    const sheet = wb.addChartSheet(bound(wb));
    const history = new EditHistory(wb);
    history.execute(TableCommands.renameTable(wb, 'Data', wb.findSheet('Data')!.tables[0]!, 'Revenue'));
    expect(sheet.chart!.source!.table).toBe('Revenue');
    expect(ChartDataResolver.resolve(wb, sheet.chart!).series[0]!.values).toEqual([10, 20, 30]);
    history.undo();
    expect(sheet.chart!.source!.table).toBe('Sales');
  });
});

describe('Range ↔ Table', () => {
  it('turns a binding into plain ranges', () => {
    const wb = book();
    const ranges = ChartTableBinding.toRanges(wb, bound(wb));
    expect(ranges.source).toBeNull();
    expect(ranges.series[0]!.valuesReference).toBe('Data.B2:B4');
  });

  it('binds ranges that are whole columns of one table', () => {
    const wb = book();
    const chart = createChart({
      categoriesReference: 'Data.A2:A4',
      series: [createSeries({ valuesReference: 'Data.C2:C4', type: 'line' })],
    });
    const result = ChartTableBinding.fromRanges(wb, chart);
    if (!result.ok) throw new Error(result.reason);
    expect(result.chart.source).toEqual({
      table: 'Sales',
      categoryColumn: 'Region',
      series: [{ column: 'Cost', type: 'line', sizeColumn: null }],
    });
  });

  it('explains why ranges that don’t line up with a table can’t be bound', () => {
    const wb = book();
    const partial = createChart({ series: [createSeries({ valuesReference: 'Data.B2:B3' })] });
    expect(ChartTableBinding.fromRanges(wb, partial)).toMatchObject({ ok: false, reason: /whole column/ });
    const outside = createChart({ series: [createSeries({ valuesReference: 'Data.F2:F4' })] });
    expect(ChartTableBinding.fromRanges(wb, outside)).toMatchObject({ ok: false, reason: /inside a table/ });
    expect(ChartTableBinding.fromRanges(wb, createChart())).toMatchObject({ ok: false });
  });
});

describe('ChartTableBinding.seed', () => {
  it('labels with the first text column and plots the number columns', () => {
    const wb = book();
    expect(ChartTableBinding.seed(wb, 'sales', 'column')).toEqual(source(['Sales', 'Cost']));
    expect(ChartTableBinding.seed(wb, 'Sales', 'pie')!.series.map((s) => s.column)).toEqual(['Sales']);
    expect(ChartTableBinding.seed(wb, 'Sales', 'scatter')).toEqual(source(['Cost'], 'Sales'));
    expect(ChartTableBinding.seed(wb, 'Nope', 'column')).toBeNull();
  });

  it('lists the table’s columns with what they hold', () => {
    expect(ChartTableBinding.columns(book(), 'Sales')).toEqual([
      { name: 'Region', column: 0, kind: 'text' },
      { name: 'Sales', column: 1, kind: 'number' },
      { name: 'Cost', column: 2, kind: 'number' },
    ]);
  });
});

describe('ODS round trip of a bound chart', () => {
  it('writes the table’s current ranges for other readers and reads the binding back', () => {
    const wb = book();
    wb.addChartSheet(bound(wb, ['Sales', 'Cost']));
    const sheet = wb.findSheet('Data')!;
    sheet.setValue(4, 0, CellValue.text('West'));
    sheet.setValue(4, 1, CellValue.number(40));
    sheet.replaceTable(sheet.tables[0]!, { ...sheet.tables[0]!, range: R('A1:C5') });

    const reopened = readOdsWithPackage(writeOds(wb)).workbook;
    const chart = reopened.findSheet('Chart1')!.chart!;
    expect(chart.source).toEqual(source(['Sales', 'Cost']));
    expect(chart.categoriesReference).toBe('Data.A2:A5');
    expect(chart.series[1]!.valuesReference).toBe('Data.C2:C5');
  });

  it('writes no binding for a range chart', () => {
    const wb = book();
    wb.addChartSheet(ChartTableBinding.toRanges(wb, bound(wb)));
    const chart = readOdsWithPackage(writeOds(wb)).workbook.findSheet('Chart1')!.chart!;
    expect(chart.source).toBeNull();
  });
});
