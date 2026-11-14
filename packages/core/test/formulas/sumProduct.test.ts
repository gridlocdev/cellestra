import { describe, expect, it } from 'vitest';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { evaluate } from './formulaHarness.ts';

/** A1:B3 = 2, 3 / "x", 5 / 4, TRUE. */
function setup(sheet: Worksheet): void {
  const rows = [
    [CellValue.number(2), CellValue.number(3)],
    [CellValue.text('x'), CellValue.number(5)],
    [CellValue.number(4), CellValue.boolean(true)],
  ];
  rows.forEach((row, r) => row.forEach((v, c) => sheet.setValue(r, c, v)));
}

// LibreOffice Calc 26.2 on the same cells.
describe('SUMPRODUCT', () => {
  it.each([
    ['SUMPRODUCT(A1:A3, B1:B3)', CellValue.number(10)],
    ['SUMPRODUCT({1;2}, {3;4})', CellValue.number(11)],
    ['SUMPRODUCT(A1:A3)', CellValue.number(6)],
    ['SUMPRODUCT(A1:B3)', CellValue.number(15)],
    ['SUMPRODUCT(3, 4)', CellValue.number(12)],
  ])('%s', (formula, expected) => {
    expect(evaluate(formula, setup)).toEqual(expected);
  });

  it.each([
    ['SUMPRODUCT(A1:A2, B1:B3)', '#VALUE!'],
    ['SUMPRODUCT(1/0, A1)', '#DIV/0!'],
  ])('%s is %s', (formula, code) => {
    const value = evaluate(formula, setup);
    expect(value.kind === 'error' ? value.error.token : value).toBe(code);
  });
});
