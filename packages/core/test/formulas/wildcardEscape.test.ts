import { describe, expect, it } from 'vitest';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { num } from './formulaHarness.ts';

/** D1:D4 = a*b, axb, a?c, abc. */
function setup(sheet: Worksheet): void {
  ['a*b', 'axb', 'a?c', 'abc'].forEach((t, r) => sheet.setValue(r, 3, CellValue.text(t)));
}

// Results LibreOffice Calc 26.2 gives: ~ makes the next * or ? a plain character.
describe('~ escapes a wildcard', () => {
  it.each([
    ['COUNTIF(D1:D4, "a~*b")', 1],
    ['COUNTIF(D1:D4, "a*b")', 2],
    ['COUNTIF(D1:D4, "<>a~*b")', 3],
    ['MATCH("a~*b", D1:D4, 0)', 1],
    ['MATCH("a~?c", D1:D4, 0)', 3],
    ['MATCH("a?c", D1:D4, 0)', 3],
  ])('%s = %s', (formula, expected) => {
    expect(num(formula, setup)).toBe(expected);
  });
});
