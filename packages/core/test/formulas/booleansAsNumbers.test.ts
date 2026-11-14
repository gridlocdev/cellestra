import { describe, expect, it } from 'vitest';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { num } from './formulaHarness.ts';

/** C1:C5 = TRUE, 1, FALSE, 0, 2. */
function setup(sheet: Worksheet): void {
  const values = [
    CellValue.boolean(true),
    CellValue.number(1),
    CellValue.boolean(false),
    CellValue.number(0),
    CellValue.number(2),
  ];
  values.forEach((v, r) => sheet.setValue(r, 2, v));
}

// Results LibreOffice Calc 26.2 gives for the same formulas: booleans are the numbers 1 and 0.
describe('Criteria and lookups treat TRUE and FALSE as 1 and 0', () => {
  it.each([
    ['COUNTIF(C1:C5, TRUE)', 2],
    ['COUNTIF(C1:C5, 1)', 2],
    ['COUNTIF(C1:C5, 0)', 2],
    ['COUNTIF(C1:C5, FALSE)', 2],
    ['COUNTIF(C1:C5, ">0")', 3],
    ['COUNTIF(C1:C5, "TRUE")', 2],
    ['COUNTIF(C1:C5, "=TRUE")', 2],
    ['COUNTIF(C1:C5, "1")', 2],
    ['COUNTIF(C1:C5, "<>1")', 3],
    ['SUMIF(C1:C5, TRUE, C1:C5)', 2],
    ['MATCH(TRUE, C1:C5, 0)', 1],
    ['MATCH(1, C1:C5, 0)', 1],
    ['MATCH(0, C1:C5, 0)', 3],
    ['MATCH(FALSE, C1:C5, 0)', 3],
  ])('%s = %s', (formula, expected) => {
    expect(num(formula, setup)).toBe(expected);
  });
});

/** A1:A3 = TRUE, FALSE, 5. */
function aggregates(sheet: Worksheet): void {
  [CellValue.boolean(true), CellValue.boolean(false), CellValue.number(5)].forEach((v, r) =>
    sheet.setValue(r, 0, v),
  );
}

// LibreOffice Calc 26.2 on the same cells: a boolean cell in a range is a number to every aggregate.
describe('Aggregates read TRUE and FALSE cells as 1 and 0', () => {
  it.each([
    ['COUNT(A1:A3)', 3],
    ['SUM(A1:A3)', 6],
    ['AVERAGE(A1:A3)', 2],
    ['MIN(A1:A3)', 0],
    ['MAX(A1:A2)', 1],
    ['PRODUCT(A1:A3)', 0],
    ['SUBTOTAL(9, A1:A3)', 6],
    ['SUBTOTAL(2, A1:A3)', 3],
  ])('%s = %s', (formula, expected) => {
    expect(num(formula, aggregates)).toBe(expected);
  });
});
