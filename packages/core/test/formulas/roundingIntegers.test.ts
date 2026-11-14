import { describe, expect, it } from 'vitest';
import { num } from './formulaHarness.ts';

describe('Rounding large integers and binary noise, as in Calc', () => {
  it.each([
    ['ROUND(1234567890123456, 0)', 1234567890123456],
    ['ROUND(1234567890123456, 2)', 1234567890123456],
    ['ROUND(-1234567890123456, 1)', -1234567890123456],
    ['ROUNDUP(1234567890123456, 0)', 1234567890123456],
    ['ROUNDDOWN(1234567890123456, 0)', 1234567890123456],
    ['ROUND(9007199254740993, 0)', 9007199254740992], // already past exact integers: left as stored
    ['ROUND(1234567890123456, -1)', 1234567890123460], // tens are still rounded
  ])('%s keeps the integer', (formula, expected) => {
    expect(num(formula)).toBe(expected);
  });

  it.each([
    ['INT((0.7+0.1)*10)', 8], // 7.999999999999999 in binary, 8 to 15 digits
    ['INT(-(0.7+0.1)*10)', -8],
    ['INT(2.9999999999999996)', 3],
    ['INT(3.7)', 3],
    ['INT(-3.2)', -4],
    ['INT(1234567890123456)', 1234567890123456],
    ['INT(0.99999999999999)', 0], // 14 nines is a real fraction, not noise
  ])('%s = %d', (formula, expected) => {
    expect(num(formula)).toBe(expected);
  });
});
