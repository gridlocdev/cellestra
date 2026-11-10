import { describe, expect, it } from 'vitest';
import { Workbook } from '../../src/model/workbook.ts';
import { CellStyles, Fills, type CellStyle, type CellStyleMap } from '../../src/styling/cellStyle.ts';
import {
  ConditionalStyles,
  parseMapCondition,
  writeMapCondition,
} from '../../src/styling/conditionalStyles.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const red: CellStyle = { ...CellStyles.default, fill: Fills.solid(0xffff0000) };
const map = (condition: string, baseCellAddress: string | null = 'Sheet1.A1'): CellStyleMap => ({
  condition,
  applyStyleName: null,
  applyStyle: red,
  baseCellAddress,
});

function sheetWith(values: (number | string | null)[][]) {
  const wb = Workbook.createNew();
  const sheet = wb.worksheets[0]!;
  values.forEach((row, r) =>
    row.forEach((v, c) => {
      if (v !== null) sheet.setValue(r, c, typeof v === 'number' ? CellValue.number(v) : CellValue.text(v));
    }),
  );
  return { wb, cs: new ConditionalStyles(wb) };
}

describe('parseMapCondition', () => {
  it('reads the forms Calc writes and writes them back', () => {
    for (const text of [
      'cell-content()>5',
      'cell-content()<=[.B1]',
      'cell-content()!=""',
      'cell-content()="Done"',
      'cell-content-is-between(1,10)',
      'cell-content-is-not-between([.$B$1],[Sheet2.C3])',
      'is-true-formula([.A1]>[.B1])',
    ]) {
      const parsed = parseMapCondition(text);
      expect(parsed, text).not.toBeNull();
      expect(writeMapCondition(parsed!)).toBe(text);
    }
    expect(parseMapCondition('of:cell-content()>=2')).toMatchObject({ kind: 'compare', operator: '>=' });
    expect(parseMapCondition('cell-content-is-between(1;2)')).toMatchObject({ kind: 'between' });
    expect(parseMapCondition('cell-content-is-whole-number()')).toBeNull();
    expect(parseMapCondition('cell-content()>')).toBeNull();
  });
});

describe('ConditionalStyles', () => {
  it('compares a cell with a value, skipping empty cells', () => {
    const { cs } = sheetWith([[3], [7], [null]]);
    const style = { ...CellStyles.default, conditionalMaps: [map('cell-content()>5')] };
    expect(cs.resolve(style, 'Sheet1', 0, 0)).toBe(style);
    expect(cs.resolve(style, 'Sheet1', 1, 0)).toBe(red);
    const lt = { ...CellStyles.default, conditionalMaps: [map('cell-content()<5')] };
    expect(cs.resolve(lt, 'Sheet1', 2, 0)).toBe(lt);
  });

  it('compares text case-insensitively and checks between', () => {
    const { cs } = sheetWith([['done', 1, 10, 11]]);
    const eq = { ...CellStyles.default, conditionalMaps: [map('cell-content()="Done"')] };
    expect(cs.resolve(eq, 'Sheet1', 0, 0)).toBe(red);
    const between = { ...CellStyles.default, conditionalMaps: [map('cell-content-is-between(1,10)')] };
    expect([1, 2, 3].map((c) => cs.resolve(between, 'Sheet1', 0, c) === red)).toEqual([true, true, false]);
    const outside = { ...CellStyles.default, conditionalMaps: [map('cell-content-is-not-between(1,10)')] };
    expect(cs.resolve(outside, 'Sheet1', 0, 3)).toBe(red);
  });

  it('never orders text against numbers, as Calc does', () => {
    const { cs } = sheetWith([['x', 7]]);
    const gt = { ...CellStyles.default, conditionalMaps: [map('cell-content()>5')] };
    const ne = { ...CellStyles.default, conditionalMaps: [map('cell-content()!=5')] };
    const outside = { ...CellStyles.default, conditionalMaps: [map('cell-content-is-not-between(1,3)')] };
    expect(cs.resolve(gt, 'Sheet1', 0, 0)).toBe(gt);
    expect(cs.resolve(ne, 'Sheet1', 0, 0)).toBe(red);
    expect(cs.resolve(outside, 'Sheet1', 0, 0)).toBe(red);
    expect(cs.resolve(gt, 'Sheet1', 0, 1)).toBe(red);
  });

  it('moves relative references with the cell, from the base cell', () => {
    // Column A is red where it is bigger than column B on the same row.
    const { cs } = sheetWith([
      [5, 1],
      [1, 5],
    ]);
    const style = { ...CellStyles.default, conditionalMaps: [map('cell-content()>[.B1]')] };
    expect(cs.resolve(style, 'Sheet1', 0, 0)).toBe(red);
    expect(cs.resolve(style, 'Sheet1', 1, 0)).toBe(style);
    const absolute = { ...CellStyles.default, conditionalMaps: [map('cell-content()>[.$B$1]')] };
    expect(cs.resolve(absolute, 'Sheet1', 1, 0)).toBe(absolute); // 1 > 1 is false
  });

  it('applies a true formula, and the first map that holds wins', () => {
    const { cs } = sheetWith([[2, 'x']]);
    const blue: CellStyle = { ...CellStyles.default, fill: Fills.solid(0xff0000ff) };
    const style = {
      ...CellStyles.default,
      conditionalMaps: [
        map('is-true-formula(LEN([.B1])>0)'),
        { ...map('cell-content()>0'), applyStyle: blue },
      ],
    };
    expect(cs.resolve(style, 'Sheet1', 0, 0)).toBe(red);
    const second = {
      ...style,
      conditionalMaps: [map('is-true-formula([.B1]="y")'), style.conditionalMaps[1]!],
    };
    expect(cs.resolve(second, 'Sheet1', 0, 0)).toBe(blue);
  });

  it('ignores conditions it cannot read and maps with no style', () => {
    const { cs } = sheetWith([[9]]);
    const style = {
      ...CellStyles.default,
      conditionalMaps: [
        map('cell-content-is-whole-number()'),
        { ...map('cell-content()>1'), applyStyle: null },
      ],
    };
    expect(cs.resolve(style, 'Sheet1', 0, 0)).toBe(style);
  });
});
