import { createChart, createSeries } from '@cellestra/charts';
import { describe, expect, it } from 'vitest';
import {
  CellRange,
  CellValue,
  ChartChecks,
  ChartTableBinding,
  quoteSheetIfNeeded,
  table as makeTable,
  Workbook,
} from '../../src/index.ts';

function book(): Workbook {
  const wb = Workbook.createNew();
  const s = wb.activeSheet;
  s.name = 'Data';
  ['Region', 'Sales', 'Cost'].forEach((h, c) => s.setValue(0, c, CellValue.text(h)));
  for (let r = 1; r <= 3; r++)
    [CellValue.text(`R${r}`), CellValue.number(r), CellValue.number(r * 2)].forEach((v, c) =>
      s.setValue(r, c, v),
    );
  return wb;
}

const messages = (wb: Workbook, chart: ReturnType<typeof createChart>) =>
  ChartChecks.problems(wb, chart).map((p) => [p.field, p.message]);

describe('ChartChecks.problems', () => {
  it('passes good and blank fields', () => {
    const chart = createChart({
      categoriesReference: 'Data.A2:A4',
      series: [
        createSeries({ valuesReference: 'Data.B2:B4', nameReference: 'Data.B1' }),
        createSeries({ valuesReference: '' }),
      ],
    });
    expect(messages(book(), chart)).toEqual([]);
  });

  it('explains references that can’t plot', () => {
    const wb = book();
    wb.addChartSheet(createChart());
    const chart = createChart({
      categoriesReference: 'Nope.A2:A4',
      series: [
        createSeries({ valuesReference: 'B2:', nameReference: 'Data.B1:C1' }),
        createSeries({ valuesReference: 'Data.B2:C4' }),
        createSeries({ valuesReference: 'Chart1.A1:A3' }),
      ],
    });
    expect(messages(wb, chart)).toEqual([
      [{ kind: 'categories' }, 'There is no sheet called “Nope”.'],
      [{ kind: 'nameCell', series: 0 }, 'A name cell is one cell, like Sheet1.B1.'],
      [{ kind: 'values', series: 0 }, '“B2:” isn’t a cell range. Write it like Sheet1.B2:B7.'],
      [{ kind: 'values', series: 1 }, 'Use one column or one row of cells, not a block.'],
      [{ kind: 'values', series: 2 }, '“Chart1” is a chart sheet. Use cells on a sheet with data.'],
    ]);
  });

  it('flags values that don’t match the labels or X values one for one', () => {
    const wb = book();
    const column = createChart({
      categoriesReference: 'Data.A2:A4',
      series: [createSeries({ valuesReference: 'Data.B2:B3' })],
    });
    expect(messages(wb, column)).toEqual([
      [{ kind: 'values', series: 0 }, 'These 2 values don’t match the 3 category labels one for one.'],
    ]);
    const scatter = createChart({
      type: 'scatter',
      series: [createSeries({ valuesReference: 'Data.C2:C4', xValuesReference: 'Data.B2:B3' })],
    });
    expect(messages(wb, scatter)).toEqual([
      [{ kind: 'values', series: 0 }, 'These 3 Y values don’t match the 2 X values one for one.'],
    ]);
  });
});

describe('ChartChecks.summary', () => {
  it('names the table and columns, or the cells', () => {
    const wb = book();
    expect(ChartChecks.summary(wb, createChart())).toBe('No data yet.');
    const ranges = createChart({
      categoriesReference: 'Data.A2:A4',
      series: [createSeries({ valuesReference: 'Data.C2:C4' })],
    });
    expect(ChartChecks.summary(wb, ranges)).toBe('From cells Data.A2:C4.');
    wb.findSheet('Data')!.addTable(makeTable('Sales', CellRange.parse('A1:C4')));
    const bound = ChartTableBinding.setTable(wb, createChart(), 'Sales');
    expect(ChartChecks.summary(wb, bound)).toBe('From table “Sales”: Sales and Cost by Region.');
    const scatter = ChartTableBinding.retype(wb, bound, 'scatter');
    expect(ChartChecks.summary(wb, scatter)).toBe('From table “Sales”: Cost against Sales.');
    wb.findSheet('Data')!.removeTable(wb.findSheet('Data')!.tables[0]!);
    expect(ChartChecks.summary(wb, bound)).toBe('From table “Sales”, which no longer exists.');
  });
});

describe('ChartChecks.sample', () => {
  it('shows the first few cells a reference points at', () => {
    const wb = Workbook.createNew();
    const s = wb.activeSheet;
    ['Jan', 'Feb', 'Mar', 'Apr'].forEach((m, r) => s.setValue(r, 0, CellValue.text(m)));
    [1, 2.5].forEach((n, c) => s.setValue(0, c + 1, CellValue.number(n)));
    const name = quoteSheetIfNeeded(s.name);
    expect(ChartChecks.sample(wb, `${name}.A1:A2`)).toBe('Jan, Feb');
    expect(ChartChecks.sample(wb, 'A1:A4')).toBe('Jan, Feb, Mar … 4 cells');
    expect(ChartChecks.sample(wb, `${name}.B1:C1`)).toBe('1, 2.5');
    expect(ChartChecks.sample(wb, `${name}.D1:D2`)).toBe('Empty cells');
    expect(ChartChecks.sample(wb, `${name}.A1:B2`)).toBeNull();
    expect(ChartChecks.sample(wb, 'Nowhere.A1')).toBeNull();
    expect(ChartChecks.sample(wb, '')).toBeNull();
  });
});
