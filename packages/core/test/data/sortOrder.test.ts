import { describe, expect, it } from 'vitest';
import { sortSpec, SortEngine } from '../../src/data/sortEngine.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Worksheet } from '../../src/model/worksheet.ts';
import { CellError } from '../../src/values/cellError.ts';
import { CellValue } from '../../src/values/cellValue.ts';

function column(values: readonly (CellValue | null)[]): Worksheet {
  const sheet = new Worksheet('S');
  values.forEach((v, row) => {
    if (v) sheet.setValue(row, 0, v);
  });
  return sheet;
}

function describeColumn(sheet: Worksheet, rows: number): string[] {
  const out: string[] = [];
  for (let row = 0; row < rows; row++) {
    const v = sheet.getValue(row, 0);
    out.push(v.kind === 'empty' ? '' : v.kind === 'error' ? 'err' : String(v.value));
  }
  return out;
}

const sortColumn = (sheet: Worksheet, rows: number, ascending: boolean): void =>
  SortEngine.sort(
    sheet,
    CellRange.fromBounds(0, 0, rows - 1, 0),
    sortSpec([{ column: 0, ascending }], false),
    'en-US',
  );

describe('Sort order', () => {
  it('orders numbers, text, booleans, errors, then empties', () => {
    const n = CellValue.number;
    const t = CellValue.text;
    const err = CellValue.error(CellError.value());
    const sheet = column([err, n(3), t('b'), CellValue.boolean(true), null, n(1), err, t('a'), n(2)]);
    sortColumn(sheet, 9, true);
    expect(describeColumn(sheet, 9)).toEqual(['1', '2', '3', 'a', 'b', 'true', 'err', 'err', '']);
  });

  it('reverses the kinds when descending but keeps empties last', () => {
    const n = CellValue.number;
    const err = CellValue.error(CellError.divideByZero());
    const sheet = column([n(1), err, null, n(3), CellValue.text('x'), n(2)]);
    sortColumn(sheet, 6, false);
    expect(describeColumn(sheet, 6)).toEqual(['err', 'x', '3', '2', '1', '']);
  });
});

describe('Sort collation', () => {
  it('sorts accented letters with their base letter in the locale', () => {
    const t = CellValue.text;
    const sheet = column([t('zebra'), t('éclair'), t('apple'), t('Echo')]);
    sortColumn(sheet, 4, true);
    expect(describeColumn(sheet, 4)).toEqual(['apple', 'Echo', 'éclair', 'zebra']);
  });

  it('ignores case and keeps the original order of case-only ties', () => {
    const t = CellValue.text;
    const sheet = column([t('b'), t('B'), t('a'), t('A')]);
    sortColumn(sheet, 4, true);
    expect(describeColumn(sheet, 4)).toEqual(['a', 'A', 'b', 'B']);
  });
});
