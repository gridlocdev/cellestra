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

  it('never merges blank labels', () => {
    expect(ChartGrouping.group(['', 'A', ''], [series([1, 2, 3])], 'sum')).toEqual({
      categories: ['', 'A', ''],
      values: [[1, 2, 3]],
    });
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
