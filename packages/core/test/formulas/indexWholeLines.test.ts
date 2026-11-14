import { describe, expect, it } from 'vitest';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { errorToken, num } from './formulaHarness.ts';

/** A1:A3 = 1, 2, 3; B1:B3 = 10, 20, 30; C1 = 100. */
function setup(sheet: Worksheet): void {
  [1, 2, 3].forEach((v, r) => sheet.setValue(r, 0, CellValue.number(v)));
  [10, 20, 30].forEach((v, r) => sheet.setValue(r, 1, CellValue.number(v)));
  sheet.setValue(0, 2, CellValue.number(100));
}

// Results LibreOffice Calc 26.2 gives for the same formulas.
describe('INDEX with a row or column of 0', () => {
  it.each([
    ['SUM(INDEX(A1:B3,0,2))', 60],
    ['SUM(INDEX(A1:B3,2,0))', 22],
    ['SUM(INDEX(A1:B3,2))', 22],
    ['SUM(INDEX(A1:B3,0,0))', 66],
    ['SUM(INDEX(A1:A3,0))', 6],
    ['SUM(INDEX(A1:C1,0))', 111],
    ['SUM(INDEX(A1:C1,0,2))', 10],
    ['SUM(INDEX(A1:B3,1,0))', 11],
    ['INDEX(A1:C1,2)', 10],
    ['INDEX(A1:C1,1,2)', 10],
    ['INDEX(A1:A3,2,1)', 2],
    ['INDEX(A1:B3,2,2)', 20],
  ])('%s = %s', (formula, expected) => {
    expect(num(formula, setup)).toBe(expected);
  });

  it.each([
    ['INDEX(A1:B3,4,1)', '#REF!'],
    ['INDEX(A1:A3,1,2)', '#REF!'],
  ])('%s is %s', (formula, token) => {
    expect(errorToken(formula, setup)).toBe(token);
  });
});
