import { describe, expect, it } from 'vitest';
import { FormulaWriter } from '../../src/formulas/formulaWriter.ts';
import { OpenFormulaTranslator } from '../../src/formulas/openFormulaTranslator.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { roundTrip } from '../serialization/helpers.ts';
import { errorToken, num } from './formulaHarness.ts';

const write = (formula: string): string => FormulaWriter.write(Parser.parse(formula));

/** A1:C3 holds 1..9 row by row. */
function grid(sheet: Worksheet): void {
  for (let r = 0; r < 3; r++)
    for (let c = 0; c < 3; c++) sheet.setValue(r, c, CellValue.number(r * 3 + c + 1));
}

describe('the space intersection operator', () => {
  it.each([
    ['A1:C3 B2:D4', 'A1:C3!B2:D4'],
    ['SUM(A1:C3 B2:D4)', 'SUM(A1:C3!B2:D4)'],
    ['SUM(A1:C3  \n B2:D4)', 'SUM(A1:C3!B2:D4)'],
    ['A:A 2:2', 'A:A!2:2'],
    ["'My Sheet'.A1:B2 'My Sheet'.B2:C3", "'My Sheet'.A1:B2!'My Sheet'.B2:C3"],
    ['(A1:B2) (B1:C3)', 'A1:B2!B1:C3'],
    ['SUM(A1:B2 B2:C3~C3)', 'SUM(A1:B2!B2:C3~C3)'], // as tight as !, so tighter than ~
    ['A1:C3 B2:D4 C3', 'A1:C3!B2:D4!C3'],
  ])('reads %j as %s', (formula, expected) => {
    expect(write(formula)).toBe(expected);
  });

  it.each([
    [' = SUM( A1 )', 'SUM(A1)'],
    ['A1 + B1', 'A1+B1'],
    ['SUM( A1 , B1 )', 'SUM(A1, B1)'],
    ['IF( A1 > 0 , "a b" , B1 )', 'IF(A1>0, "a b", B1)'],
    ['SUM(A1,\n  B1)', 'SUM(A1, B1)'],
    ['A1 & " " & B1', 'A1&" "&B1'],
    ['A1 %', 'A1%'],
    ['- A1', '-A1'],
    ['{1 , 2 ; 3 , 4}', '{1,2;3,4}'],
  ])('keeps other spaces meaningless: %j', (formula, expected) => {
    expect(write(formula)).toBe(expected);
  });

  it.each(['A1 1', 'A1 "x"', 'A1:B2 TRUE', '1 A1', '(A1:B2)(B1:C3)'])('refuses %j', (formula) => {
    expect(Parser.tryParse(formula).ok).toBe(false);
  });

  it.each([
    ['SUM(A1:C3 B2:D4)', 5 + 6 + 8 + 9],
    ['A1:C3 C3', 9],
    ['A:A 2:2', 4],
    ['SUM(A1:C1 B1:B3 B1)', 2],
  ])('%s = %d', (formula, expected) => {
    expect(num(formula, grid)).toBe(expected);
  });

  it('ranges that do not meet are #NULL!', () => {
    expect(errorToken('SUM(A1:A3 C1:C3)', grid)).toBe('#NULL!');
  });

  it('is written to OpenFormula as !', () => {
    expect(OpenFormulaTranslator.toOpenFormula('SUM(A1:C3 B2:D4)')).toBe('of:=SUM([.A1:.C3]![.B2:.D4])');
    expect(OpenFormulaTranslator.fromOpenFormula('of:=SUM([.A1:.C3]![.B2:.D4])')).toBe('SUM(A1:C3!B2:D4)');
  });

  it('round-trips through an .ods file', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.setFormula(5, 0, 'SUM(A1:C3 B2:D4)');
    expect(roundTrip(workbook).activeSheet.find(5, 0)?.formula).toBe('SUM(A1:C3!B2:D4)');
  });
});
