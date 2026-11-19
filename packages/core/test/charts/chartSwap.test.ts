import { createChart, createSeries } from '@cellestra/charts';
import { describe, expect, it } from 'vitest';
import { ChartSwap, Workbook } from '../../src/index.ts';

function book(): Workbook {
  const wb = Workbook.createNew();
  wb.activeSheet.name = 'Data';
  return wb;
}

const columns = createChart({
  title: 'Spend',
  categoriesReference: 'Data.A2:A4',
  series: [
    createSeries({ valuesReference: 'Data.B2:B4', nameReference: 'Data.B1', type: 'line' }),
    createSeries({ valuesReference: 'Data.C2:C4', nameReference: 'Data.C1' }),
  ],
});

describe('ChartSwap.swap', () => {
  it('turns column series into row series, labels into names and names into labels', () => {
    const result = ChartSwap.swap(book(), columns);
    if (!result.ok) throw new Error(result.reason);
    expect(result.chart.title).toBe('Spend');
    expect(result.chart.categoriesReference).toBe('Data.B1:C1');
    expect(result.chart.series.map((s) => [s.valuesReference, s.nameReference, s.type])).toEqual([
      ['Data.B2:C2', 'Data.A2', null],
      ['Data.B3:C3', 'Data.A3', null],
      ['Data.B4:C4', 'Data.A4', null],
    ]);
  });

  it('gives the chart back when swapped twice', () => {
    const wb = book();
    const once = ChartSwap.swap(wb, columns);
    if (!once.ok) throw new Error(once.reason);
    const twice = ChartSwap.swap(wb, once.chart);
    if (!twice.ok) throw new Error(twice.reason);
    expect(twice.chart.categoriesReference).toBe(columns.categoriesReference);
    expect(twice.chart.series.map((s) => [s.valuesReference, s.nameReference])).toEqual([
      ['Data.B2:B4', 'Data.B1'],
      ['Data.C2:C4', 'Data.C1'],
    ]);
  });

  it('works without labels or names', () => {
    const plain = createChart({ series: [createSeries({ valuesReference: 'Data.B2:B3' })] });
    const result = ChartSwap.swap(book(), plain);
    if (!result.ok) throw new Error(result.reason);
    expect(result.chart.categoriesReference).toBeNull();
    expect(result.chart.series.map((s) => [s.valuesReference, s.nameReference])).toEqual([
      ['Data.B2', null],
      ['Data.B3', null],
    ]);
  });

  it('refuses charts it can’t swap, saying why', () => {
    const wb = book();
    const apart = createChart({
      series: [
        createSeries({ valuesReference: 'Data.B2:B4' }),
        createSeries({ valuesReference: 'Data.D2:D4' }),
      ],
    });
    expect(ChartSwap.swap(wb, apart)).toMatchObject({ ok: false, reason: /side by side/ });
    expect(ChartSwap.swap(wb, { ...columns, type: 'scatter' })).toMatchObject({
      ok: false,
      reason: /X against Y/,
    });
    expect(ChartSwap.swap(wb, createChart())).toMatchObject({ ok: false });
    const misaligned = { ...columns, categoriesReference: 'Data.A1:A4' };
    expect(ChartSwap.swap(wb, misaligned)).toMatchObject({ ok: false, reason: /line up/ });
  });
});
