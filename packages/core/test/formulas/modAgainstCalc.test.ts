import { describe, expect, it } from 'vitest';
import { num } from './formulaHarness.ts';

// Results LibreOffice Calc 26.2 gives for the same formulas.
describe('MOD matches Calc where binary noise decides the remainder', () => {
  it.each([
    ['MOD(0.3, 0.1)', 0],
    ['MOD(1E15+0.5, 1)', 0],
    ['MOD(1E20, 3)', 0],
    ['MOD(2^60, 7)', 0],
    ['MOD(5.1, 1)', 0.09999999999999964],
    ['MOD(10, 3)', 1],
  ])('%s = %s', (formula, expected) => {
    expect(num(formula)).toBe(expected);
  });
});
