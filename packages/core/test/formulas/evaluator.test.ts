import { describe, expect, it } from 'vitest';
import { Evaluator } from '../../src/formulas/evaluation/evaluator.ts';
import { WorkbookEvaluationContext } from '../../src/formulas/evaluation/workbookEvaluationContext.ts';
import { FunctionRegistry } from '../../src/formulas/functions/functionRegistry.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { bool, errorToken, INVARIANT, num, text } from './formulaHarness.ts';

describe('Evaluator', () => {
  it.each([
    ['1+2', 3],
    ['2*3', 6],
    ['10-4', 6],
    ['10/4', 2.5],
    ['2^10', 1024],
    ['2^3^2', 64], // left-associative
    ['-5', -5],
    ['-2^2', 4], // a sign binds tighter than ^, as in Calc and Excel: (-2)^2
    ['-(2^2)', -4], // explicit parens flip it back
    ['-2^2*2', 8], // = ((-2)^2)*2
    ['2*-3^2', 18], // = 2*((-3)^2), as in Calc
    ['-2^-2', 0.25], // = (-2)^(-2)
    ['-2%^2', 0.0004], // % binds tighter still: (-(2%))^2
    ['50%', 0.5],
    ['(1+2)*3', 9],
    ['"5"+3', 8], // text coerces in arithmetic
  ])('arithmetic %s = %d', (formula, expected) => {
    expect(num(formula)).toBeCloseTo(expected, 10);
  });

  it('division by zero is an error', () => {
    expect(errorToken('10/0')).toBe('#DIV/0!');
  });

  it.each([
    ['1<2', true],
    ['2<=2', true],
    ['3>4', false],
    ['1=1', true],
    ['2<>3', true],
    ['"a"="A"', true], // text equality is case-insensitive
    ['"apple"<"banana"', true],
  ])('comparison %s is %s', (formula, expected) => {
    expect(bool(formula)).toBe(expected);
  });

  it('concatenation', () => {
    expect(text('"a"&"b"')).toBe('ab');
    expect(text('1&2')).toBe('12');
  });

  it('cell reference reads value', () => {
    const result = num('A1+A2', (s) => {
      s.setValue(0, 0, CellValue.number(10));
      s.setValue(1, 0, CellValue.number(5));
    });
    expect(result).toBe(15);
  });

  it('blank cell is zero in arithmetic', () => {
    expect(num('A1+5')).toBe(5);
  });

  it('bare range in scalar position is value error', () => {
    expect(errorToken('A1:A3+1')).toBe('#VALUE!');
  });

  it('unknown name is name error', () => {
    expect(errorToken('NotDefined')).toBe('#NAME?');
  });

  it('unknown function is name error', () => {
    expect(errorToken('BOGUS(1)')).toBe('#NAME?');
  });

  it('errors propagate through arithmetic', () => {
    expect(errorToken('(1/0)+5')).toBe('#DIV/0!');
  });

  it('named range resolves to single cell', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.number(0.2));
    workbook.addNamedRange({ name: 'TaxRate', target: CellRange.parse('A1') });

    const context = new WorkbookEvaluationContext(workbook, sheet.name, INVARIANT);
    const evaluator = new Evaluator(context, FunctionRegistry.createDefault());
    const value = evaluator.evaluate(Parser.parse('TaxRate*100'));

    expect(CellValue.asNumber(value)).toBeCloseTo(20, 10);
  });
});
