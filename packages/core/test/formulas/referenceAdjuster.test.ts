import { describe, expect, it } from 'vitest';
import type { Expr } from '../../src/formulas/ast/expr.ts';
import { FormulaWriter } from '../../src/formulas/formulaWriter.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';
import { ReferenceAdjuster } from '../../src/formulas/references/referenceAdjuster.ts';

const write = (expr: Expr): string => FormulaWriter.write(expr);
const parse = (formula: string): Expr => Parser.parse(formula);

describe('ReferenceAdjuster', () => {
  it.each([
    ['A1+B1', 'B2+C2'],
    ['$A$1+B1', '$A$1+C2'],
    ['$A1+A$1', '$A2+B$1'],
    ['SUM(A1:A3)', 'SUM(B2:B4)'],
  ])('copy shifts relative and keeps absolute: %s', (formula, expected) => {
    expect(write(ReferenceAdjuster.forCopy(parse(formula), 1, 1))).toBe(expected);
  });

  it('insert rows pushes later references down', () => {
    expect(write(ReferenceAdjuster.forInsertRows(parse('A5'), 'Sheet1', 'Sheet1', 2, 2))).toBe('A7');
  });

  it('insert rows ignores other sheets', () => {
    expect(write(ReferenceAdjuster.forInsertRows(parse('Sheet2.A5'), 'Sheet1', 'Sheet1', 2, 2))).toBe(
      'Sheet2.A5',
    );
  });

  it('delete rows breaks a reference to a deleted cell', () => {
    expect(write(ReferenceAdjuster.forDeleteRows(parse('A3'), 'Sheet1', 'Sheet1', 2, 1))).toBe('#REF!');
  });

  it('delete rows shrinks a spanning range', () => {
    expect(write(ReferenceAdjuster.forDeleteRows(parse('SUM(A1:A5)'), 'Sheet1', 'Sheet1', 1, 2))).toBe(
      'SUM(A1:A3)',
    );
  });

  it('delete sheet invalidates cross-sheet references', () => {
    expect(write(ReferenceAdjuster.forDeleteSheet(parse('Data.A1+1'), 'Data', 'Sheet1'))).toBe('#REF!+1');
  });

  it('delete sheet leaves other references untouched', () => {
    expect(write(ReferenceAdjuster.forDeleteSheet(parse('A1+Other.B2'), 'Data', 'Sheet1'))).toBe(
      'A1+Other.B2',
    );
  });

  it('insert columns shifts references right', () => {
    expect(write(ReferenceAdjuster.forInsertColumns(parse('C1'), 'Sheet1', 'Sheet1', 1, 1))).toBe('D1');
  });

  it('insert inside a range grows it', () => {
    // Inserting a row at the last covered row (index 4 = A5) grows the range.
    expect(write(ReferenceAdjuster.forInsertRows(parse('SUM(A1:A5)'), 'Sheet1', 'Sheet1', 4, 1))).toBe(
      'SUM(A1:A6)',
    );
  });

  it('insert just past a range does not grow it', () => {
    expect(write(ReferenceAdjuster.forInsertRows(parse('SUM(A1:A5)'), 'Sheet1', 'Sheet1', 5, 1))).toBe(
      'SUM(A1:A5)',
    );
  });

  it('insert pushing a reference past the last row yields #REF!', () => {
    expect(write(ReferenceAdjuster.forInsertRows(parse('A1048576'), 'Sheet1', 'Sheet1', 0, 1))).toBe('#REF!');
  });

  it('insert pushing a reference past the last column yields #REF!', () => {
    expect(write(ReferenceAdjuster.forInsertColumns(parse('XFD1'), 'Sheet1', 'Sheet1', 0, 1))).toBe('#REF!');
  });

  it("insert pushing only a range's far corner off the sheet clamps it", () => {
    // Both corners land on the last row, collapsing the range to a single cell.
    expect(
      write(ReferenceAdjuster.forInsertRows(parse('SUM(A1048575:A1048576)'), 'Sheet1', 'Sheet1', 0, 1)),
    ).toBe('SUM(A1048576)');
  });

  it('insert pushing a whole range off the sheet yields #REF!', () => {
    expect(
      write(ReferenceAdjuster.forInsertRows(parse('SUM(A1048576:B1048576)'), 'Sheet1', 'Sheet1', 0, 1)),
    ).toBe('SUM(#REF!)');
  });
});

describe('FormulaWriter', () => {
  it.each([
    ['1+2*3', '1+2*3'],
    ['(1+2)*3', '(1+2)*3'],
    ['2^3^2', '2^3^2'],
    ['1-(2-3)', '1-(2-3)'],
    ['1-2+3', '1-2+3'],
    ['-A1%', '-A1%'],
    ['-2^2', '-2^2'], // a sign binds tighter than ^: this is (-2)^2
    ['(-2)^2', '-2^2'], // the same tree, so the parentheses are redundant
    ['-(2^2)', '-(2^2)'], // parens preserved — changes meaning
    ['2^-3', '2^-3'], // sign on the exponent
    ['-A1%^2', '-A1%^2'],
    ['(-A1)%', '(-A1)%'],
    ['-2*3', '-2*3'],
  ])('round-trips %s with minimal parens', (formula, expected) => {
    expect(write(parse(formula))).toBe(expected);
  });

  it('writes string literals with doubled quotes', () => {
    expect(write(parse('"a""b"'))).toBe('"a""b"');
  });
});
