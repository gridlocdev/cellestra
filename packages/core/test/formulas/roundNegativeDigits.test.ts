import { describe, expect, it } from 'vitest';
import { num } from './formulaHarness.ts';

describe('Rounding to tens, hundreds and beyond', () => {
  it.each([
    ['ROUND(123456, -4)', 120000],
    ['ROUND(55555, -4)', 60000],
    ['ROUND(987654.321, -4)', 990000],
    ['ROUND(123456, -5)', 100000],
    ['ROUND(-55555, -5)', -100000],
    ['ROUNDUP(123456, -4)', 130000],
    ['ROUNDDOWN(987654.321, -5)', 900000],
    ['ROUND(1234.5678, -2)', 1200],
    ['ROUND(1.005, 2)', 1.01],
  ])('%s = %d exactly', (formula, expected) => {
    expect(num(formula)).toBe(expected);
  });
});
