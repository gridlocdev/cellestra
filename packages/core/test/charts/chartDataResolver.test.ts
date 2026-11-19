import { createChart, createSeries } from '@cellestra/charts';
import { describe, expect, it } from 'vitest';
import { ChartDataResolver } from '../../src/charts/chartDataResolver.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

function buildData(): Workbook {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet; // Sheet1
  sheet.setValue(0, 1, CellValue.text('Sales')); // B1 header / series name
  sheet.setValue(0, 0, CellValue.text('Category')); // A1
  const labels = ['Jan', 'Feb', 'Mar'];
  const values = [5, 8, 3];
  labels.forEach((label, i) => {
    sheet.setValue(i + 1, 0, CellValue.text(label));
    sheet.setValue(i + 1, 1, CellValue.number(values[i]!));
  });
  return workbook;
}

describe('ChartDataResolver', () => {
  it('resolves categories, values and the series name from a cell', () => {
    const chart = createChart({
      categoriesReference: 'Sheet1.A2:A4',
      series: [createSeries({ nameReference: 'Sheet1.B1', valuesReference: 'Sheet1.B2:B4' })],
    });

    const resolved = ChartDataResolver.resolve(buildData(), chart);

    expect(resolved.categories).toEqual(['Jan', 'Feb', 'Mar']);
    expect(resolved.series).toHaveLength(1);
    expect(resolved.series[0]!.name).toBe('Sales');
    expect(resolved.series[0]!.values).toEqual([5, 8, 3]);
    expect(resolved.isEmpty).toBe(false);
  });

  it('plots whole columns up to the last used row', () => {
    const chart = createChart({
      categoriesReference: 'Sheet1.A:A',
      series: [createSeries({ valuesReference: 'Sheet1.B:B' })],
    });
    const resolved = ChartDataResolver.resolve(buildData(), chart);
    expect(resolved.categories).toEqual(['Category', 'Jan', 'Feb', 'Mar']);
    expect(resolved.series[0]!.values).toEqual([Number.NaN, 5, 8, 3]);
  });

  it('an unqualified reference defaults to the first data sheet', () => {
    const chart = createChart({ series: [createSeries({ valuesReference: 'B2:B4' })] }); // no sheet qualifier

    const resolved = ChartDataResolver.resolve(buildData(), chart);

    expect(resolved.series[0]!.values).toEqual([5, 8, 3]);
    expect(resolved.series[0]!.name).toBe('Series 1'); // no name reference / literal → fallback
  });

  it('blank or text cells become NaN gaps', () => {
    const workbook = buildData();
    workbook.activeSheet.setValue(2, 1, CellValue.empty); // blank out B3
    const chart = createChart({ series: [createSeries({ valuesReference: 'Sheet1.B2:B4' })] });

    const values = ChartDataResolver.resolve(workbook, chart).series[0]!.values;

    expect(values[0]).toBe(5);
    expect(values[1]).toBeNaN();
    expect(values[2]).toBe(3);
  });

  it('a chart with no series is empty', () => {
    expect(ChartDataResolver.resolve(buildData(), createChart()).isEmpty).toBe(true);
  });

  it('numeric category labels use invariant 0.###### formatting', () => {
    const workbook = buildData();
    workbook.activeSheet.setValue(1, 0, CellValue.number(2.5));
    workbook.activeSheet.setValue(2, 0, CellValue.number(1 / 3));
    const chart = createChart({
      categoriesReference: 'A2:A4',
      series: [createSeries({ valuesReference: 'B2:B4' })],
    });

    expect(ChartDataResolver.resolve(workbook, chart).categories).toEqual(['2.5', '0.333333', 'Mar']);
  });
});
