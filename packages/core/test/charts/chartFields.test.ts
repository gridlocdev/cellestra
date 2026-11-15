import { createChart, createSeries } from '@cellestra/charts';
import { describe, expect, it } from 'vitest';
import { CellRange, ChartFields, Workbook } from '../../src/index.ts';

describe('ChartFields', () => {
  it('reads and writes a chart’s reference fields by name', () => {
    const chart = createChart({ series: [createSeries({ valuesReference: 'Sheet1.B2:B4' })] });
    expect(ChartFields.get(chart, { kind: 'values', series: 0 })).toBe('Sheet1.B2:B4');
    expect(ChartFields.get(chart, { kind: 'categories' })).toBeNull();
    const next = ChartFields.set(chart, { kind: 'categories' }, 'Sheet1.A2:A4');
    expect(next.categoriesReference).toBe('Sheet1.A2:A4');
    expect(chart.categoriesReference).toBeNull();
    const added = ChartFields.set(chart, { kind: 'values', series: 1 }, 'Sheet1.C2:C4');
    expect(added.series.map((s) => s.valuesReference)).toEqual(['Sheet1.B2:B4', 'Sheet1.C2:C4']);
    expect(
      ChartFields.set(chart, { kind: 'nameCell', series: 0 }, 'Sheet1.B1').series[0]!.nameReference,
    ).toBe('Sheet1.B1');
  });

  it('writes a picked range the way charts store references', () => {
    expect(ChartFields.reference('My Data', CellRange.parse('$A$2:B5'))).toBe("'My Data'.A2:B5");
    expect(ChartFields.label({ kind: 'xValues', series: 1 })).toBe('series 2 X values');
  });

  it('picks on the sheet the field points at, else the chart’s data sheet, else the first sheet', () => {
    const wb = Workbook.createNew();
    wb.addSheet('Data');
    const chart = createChart({ series: [createSeries({ valuesReference: 'Data.B2:B4' })] });
    const own = ChartFields.pickTarget(wb, chart, { kind: 'values', series: 0 });
    expect(own?.sheet).toBe('Data');
    expect(own?.range?.toString()).toBe('B2:B4');
    expect(ChartFields.pickTarget(wb, chart, { kind: 'categories' })).toEqual({ sheet: 'Data', range: null });
    expect(ChartFields.pickTarget(wb, createChart(), { kind: 'categories' })).toEqual({
      sheet: 'Sheet1',
      range: null,
    });
  });
});
