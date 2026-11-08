import { describe, expect, it } from 'vitest';
import { approxAdd, approxEqual, approxSub } from '../../src/formulas/evaluation/valueComparer.ts';
import { KahanSum } from '../../src/formulas/functions/kahanSum.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { bool, num } from './formulaHarness.ts';

/** Fill cells from a map of A1-style column letter to the values going down from row 1. */
const columns =
  (data: Record<string, readonly (number | string)[]>) =>
  (sheet: Worksheet): void => {
    for (const [letter, values] of Object.entries(data)) {
      values.forEach((v, row) =>
        sheet.setValue(
          row,
          letter.charCodeAt(0) - 65,
          typeof v === 'number' ? CellValue.number(v) : CellValue.text(v),
        ),
      );
    }
  };

const sum = (values: readonly number[]): number => {
  const k = new KahanSum();
  for (const v of values) k.add(v);
  return k.get();
};

// Expected values below were measured in LibreOffice Calc 26.2 (headless, .fods → CSV).
describe('+ and − cancel like rtl::math::approxAdd / approxSub', () => {
  it('a sum that cancels within approxEqual is exactly 0', () => {
    expect(num('0.1+0.2-0.3')).toBe(0);
    expect(bool('0.1+0.2-0.3=0')).toBe(true);
    expect(num('0.3-0.1-0.2')).toBe(0);
    expect(num('-0.3+0.1+0.2')).toBe(0);
    expect(num('-0.1-0.2+0.3')).toBe(0);
    expect(num('0.1*3-0.3')).toBe(0);
    expect(num('1+1E-16-1')).toBe(0);
    expect(num('0.2+0.5+1E15-1E15')).toBe(0);
  });

  it('ordinary arithmetic is unchanged', () => {
    expect(num('0.1+0.2')).toBe(0.1 + 0.2);
    expect(num('2.2-2.1')).toBe(2.2 - 2.1);
    expect(num('1-1E-15')).toBe(1 - 1e-15);
    expect(num('1+2')).toBe(3);
    expect(num('10-3.5')).toBe(6.5);
    // Two different integers a double holds exactly are never approximately equal.
    expect(num('1E15+1-1E15')).toBe(1);
  });

  it('the helpers snap only when the signs allow cancellation', () => {
    expect(approxAdd(0.30000000000000004, -0.3)).toBe(0);
    expect(approxAdd(0.30000000000000004, 0.3)).toBe(0.30000000000000004 + 0.3);
    expect(approxSub(0.30000000000000004, 0.3)).toBe(0);
    expect(approxSub(0.30000000000000004, -0.3)).toBe(0.30000000000000004 + 0.3);
    expect(approxSub(-0.30000000000000004, -0.3)).toBe(0);
  });
});

describe('approxEqual follows rtl_math_approxEqual', () => {
  it('allows half the 15th significant digit as well as 2^-48', () => {
    expect(bool('1=1+4E-15')).toBe(true);
    expect(approxEqual(1, 1 + 6e-15)).toBe(false);
  });

  it('keeps distinct exact integers apart and opposite signs apart', () => {
    expect(bool('1E15=1E15+1')).toBe(false);
    expect(approxEqual(1e15, 1e15 + 0.5)).toBe(true);
    expect(approxEqual(-1e-300, 1e-300)).toBe(false);
    expect(approxEqual(Infinity, Infinity)).toBe(true);
    expect(approxEqual(Number.MAX_VALUE, Infinity)).toBe(false);
  });
});

describe('KahanSum', () => {
  it('snaps a total that the last value cancels to 0', () => {
    expect(sum([0.1, 0.2, -0.3])).toBe(0);
    expect(sum([1e16, 1, -1e16])).toBe(0);
  });

  it('otherwise compensates the rounding error', () => {
    expect(sum(Array.from({ length: 10 }, () => 0.1))).toBe(1);
    expect(sum([1e16, 1, 1, -1e16, 0.5])).toBe(2.5);
    expect(sum([])).toBe(0);
    expect(sum([0, 0])).toBe(0);
    expect(sum([7])).toBe(7);
  });
});

describe('SUM, AVERAGE and SUMIF use compensated summation', () => {
  const data = columns({
    D: [0.1, 0.2, -0.3],
    E: [1e16, 1, -1e16],
    G: ['a', 'a', 'a'],
    H: [1.5, 2.25, 3],
    K: Array.from({ length: 10 }, () => 0.1),
    L: [0.2, -1e16, 1e16],
    M: [-1.1, 1e16],
    N: [0.1, 1e15],
  });

  it('cancel to exactly 0', () => {
    expect(num('SUM(0.1,0.2,-0.3)')).toBe(0);
    expect(bool('SUM(0.1,0.2,-0.3)=0')).toBe(true);
    expect(num('SUM(1E16,1,-1E16)')).toBe(0);
    expect(num('SUM(1,1E-16,-1)')).toBe(0);
    expect(num('AVERAGE(0.1,0.2,-0.3)')).toBe(0);
    expect(num('AVERAGE(D1:D3)', data)).toBe(0);
    expect(num('SUMIF(G1:G3,"a",D1:D3)', data)).toBe(0);
    expect(num('SUMIF(G1:G3,"a",E1:E3)', data)).toBe(0);
  });

  it('leave ordinary sums alone', () => {
    expect(num('SUM(1,2,3)')).toBe(6);
    expect(num('SUM(0.1,0.2)')).toBe(0.1 + 0.2);
    expect(num('SUM(H1:H3)', data)).toBe(6.75);
    expect(num('AVERAGE(H1:H3)', data)).toBe(2.25);
    expect(num('SUM(K1:K10)', data)).toBe(1);
    expect(bool('SUM(K1:K10)=1', data)).toBe(true);
  });

  it('add the arguments last to first, as Calc does', () => {
    expect(num('SUM(0.2,-1E16,1E16)')).toBe(0.2);
    expect(num('SUM(1E16,-1E16,0.2)')).toBe(0);
    expect(num('SUM(1E15,-1E15,1E-16)')).toBe(0);
    expect(num('SUM(0.2,0.5,1E15,-1E15)')).toBe(0.7);
    expect(num('AVERAGE(0.2,-1E16,1E16)*3')).toBeCloseTo(0.2, 15);
  });

  it('walk a range column by column, top to bottom', () => {
    expect(num('AVERAGE(L1:L3)*3', data)).toBe(0);
    expect(num('SUMIF(L1:L3,"<>0")', data)).toBe(0);
    // Row by row would give 1.1E16; Calc's column order gives 2 less.
    expect(num('SUM(M1:N2)', data)).toBe(1.1e16 - 2);
    expect(num('SUMIF(M1:N2,"<>0")', data)).toBe(1.1e16 - 2);
    expect(num('AVERAGE(M1:N2)*4', data)).toBe(1.1e16 - 2);
  });
});
