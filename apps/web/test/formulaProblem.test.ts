import { CellError, CellValue } from '@cellestra/core';
import { describe, expect, it } from 'vitest';
import { formulaProblem } from '../src/app/controller.ts';

describe('formulaProblem', () => {
  it('says a formula that does not parse is incomplete, and why', () => {
    const text = formulaProblem({ formula: 'SUM(A1:', value: CellValue.error(CellError.value('x')) });
    expect(text).toMatch(/^This formula isn’t complete: \S/);
  });

  it('gives the reason for #NAME?', () => {
    const value = CellValue.error(CellError.name("Unknown function 'NOPE'."));
    expect(formulaProblem({ formula: 'NOPE(1)', value })).toBe("#NAME?: Unknown function 'NOPE'.");
  });

  it('says nothing for values, working formulas and errors that come from the data', () => {
    expect(formulaProblem(undefined)).toBeNull();
    expect(formulaProblem({ formula: null, value: CellValue.number(1) })).toBeNull();
    expect(formulaProblem({ formula: 'A1+1', value: CellValue.number(2) })).toBeNull();
    expect(
      formulaProblem({ formula: '1/0', value: CellValue.error(CellError.divideByZero('x')) }),
    ).toBeNull();
  });
});
