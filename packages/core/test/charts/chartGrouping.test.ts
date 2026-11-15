import { createChart, createSeries, type ChartAggregate } from '@cellestra/charts';
import { describe, expect, it } from 'vitest';
import {
  CellValue,
  ChartDataResolver,
  ChartGrouping,
  readOdsWithPackage,
  Workbook,
  writeOds,
} from '../../src/index.ts';

const series = (values: number[], filled = values.map((v) => !Number.isNaN(v))) => ({ values, filled });

describe('ChartGrouping.group', () => {
  const labels = ['North', 'South', 'North', 'East', 'South'];
  const sales = series([10, 20, 30, 5, NaN]);

  it.each<[ChartAggregate, number[]]>([
    ['sum', [40, 20, 5]],
    ['average', [20, 20, 5]],
    ['count', [2, 1, 1]],
    ['min', [10, 20, 5]],
    ['max', [30, 20, 5]],
  ])('combines repeated labels by %s, in first-appearance order', (aggregate, expected) => {
    const grouped = ChartGrouping.group(labels, [sales], aggregate);
    expect(grouped.categories).toEqual(['North', 'South', 'East']);
    expect(grouped.values[0]).toEqual(expected);
    expect(grouped.rowCounts).toEqual([2, 2, 1]);
  });

  it('counts text cells too, and leaves a group with no numbers a gap', () => {
    const grouped = ChartGrouping.group(
      ['A', 'A', 'B'],
      [series([NaN, NaN, NaN], [true, true, false])],
      'count',
    );
    expect(grouped.values[0]).toEqual([2, 0]);
    expect(ChartGrouping.group(['A', 'A'], [series([NaN, NaN])], 'sum').values[0]![0]).toBeNaN();
  });

  it('merges blank and whitespace labels into one "(blank)" category', () => {
    const grouped = ChartGrouping.group(['', 'A', '  ', 'B'], [series([1, 2, 3, 4])], 'sum');
    expect(grouped.categories).toEqual(['(blank)', 'A', 'B']);
    expect(grouped.values).toEqual([[4, 2, 4]]);
    expect(grouped.rowCounts).toEqual([2, 1, 1]);
  });

  it('skips fully empty rows, such as empty rows at the bottom of a table', () => {
    const grouped = ChartGrouping.group(
      ['A', 'B', '', ' ', ''],
      [series([1, 2, NaN, NaN, NaN]), series([3, NaN, NaN, NaN, NaN])],
      'sum',
    );
    expect(grouped.categories).toEqual(['A', 'B']);
    expect(grouped.values).toEqual([
      [1, 2],
      [3, NaN],
    ]);
  });

  it('merges labels that differ only in leading or trailing spaces', () => {
    const grouped = ChartGrouping.group(['North', ' North ', 'South'], [series([1, 2, 3])], 'sum');
    expect(grouped.categories).toEqual(['North', 'South']);
    expect(grouped.values).toEqual([[3, 3]]);
    expect(grouped.rowCounts).toEqual([2, 1]);
  });

  it('leaves distinct labels as they are', () => {
    const values = [1, NaN];
    const grouped = ChartGrouping.group(['A', 'B'], [series(values)], 'average');
    expect(grouped.values[0]).toBe(values);
  });
});

/** Sheet1.A1:B5: Region | Sales with North twice. */
function book(): Workbook {
  const wb = Workbook.createNew();
  const s = wb.activeSheet;
  [
    ['North', 10],
    ['South', 20],
    ['North', 30],
  ].forEach(([label, value], r) => {
    s.setValue(r, 0, CellValue.text(label as string));
    s.setValue(r, 1, CellValue.number(value as number));
  });
  return wb;
}

describe('Grouping in the resolver', () => {
  it('groups a range chart’s repeated categories by the chart’s aggregate', () => {
    const chart = createChart({
      type: 'pie',
      aggregate: 'max',
      categoriesReference: 'Sheet1.A1:A3',
      series: [createSeries({ valuesReference: 'Sheet1.B1:B3' })],
    });
    const resolved = ChartDataResolver.resolve(book(), chart);
    expect(resolved.categories).toEqual(['North', 'South']);
    expect(resolved.series[0]!.values).toEqual([30, 20]);
    expect(resolved.aggregate).toBe('max');
    expect(resolved.rowCounts).toEqual([2, 1]);
  });

  it('has no row counts when no rows merged', () => {
    const chart = createChart({
      categoriesReference: 'Sheet1.A1:A2',
      series: [createSeries({ valuesReference: 'Sheet1.B1:B2' })],
    });
    expect(ChartDataResolver.resolve(book(), chart).rowCounts).toBeNull();
  });

  it('names series uniquely in the legend', () => {
    const wb = book();
    wb.activeSheet.setValue(0, 3, CellValue.text('  '));
    const chart = createChart({
      categoriesReference: 'Sheet1.A1:A3',
      series: [
        createSeries({ valuesReference: 'Sheet1.B1:B3', nameReference: 'Sheet1.D1' }),
        createSeries({ valuesReference: 'Sheet1.B1:B3', name: 'Sales' }),
        createSeries({ valuesReference: 'Sheet1.B1:B3', name: 'Sales' }),
      ],
    });
    expect(ChartDataResolver.resolve(wb, chart).series.map((s) => s.name)).toEqual([
      'Series 1',
      'Sales',
      'Sales (2)',
    ]);
  });

  it('drops fully empty rows from a scatter chart', () => {
    const wb = book();
    wb.activeSheet.setValue(4, 1, CellValue.number(40));
    const chart = createChart({
      type: 'scatter',
      series: [createSeries({ valuesReference: 'Sheet1.B1:B5', xValuesReference: 'Sheet1.B1:B5' })],
    });
    const s = ChartDataResolver.resolve(wb, chart).series[0]!;
    expect(s.values).toEqual([10, 20, 30, 40]);
    expect(s.xValues).toEqual([10, 20, 30, 40]);
  });

  it('never groups a scatter chart', () => {
    const wb = book();
    wb.activeSheet.setValue(2, 0, CellValue.number(1));
    const chart = createChart({
      type: 'scatter',
      categoriesReference: 'Sheet1.A1:A3',
      series: [createSeries({ valuesReference: 'Sheet1.B1:B3', xValuesReference: 'Sheet1.B1:B3' })],
    });
    expect(ChartDataResolver.resolve(wb, chart).series[0]!.values).toEqual([10, 20, 30]);
  });

  it('round-trips the aggregate through the .ods, and reads a chart without one as sum', () => {
    const wb = book();
    wb.addChartSheet(
      createChart({ aggregate: 'average', series: [createSeries({ valuesReference: 'Sheet1.B1:B3' })] }),
    );
    expect(readOdsWithPackage(writeOds(wb)).workbook.findSheet('Chart1')!.chart!.aggregate).toBe('average');
    const plain = book();
    plain.addChartSheet(createChart({ series: [createSeries({ valuesReference: 'Sheet1.B1:B3' })] }));
    const bytes = writeOds(plain);
    expect(readOdsWithPackage(bytes).workbook.findSheet('Chart1')!.chart!.aggregate).toBe('sum');
  });
});
