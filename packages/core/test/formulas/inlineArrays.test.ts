import { describe, expect, it } from 'vitest';
import { FormulaSignatureScanner } from '../../src/formulas/functions/formulaSignatureScanner.ts';
import { FormulaWriter } from '../../src/formulas/formulaWriter.ts';
import { OpenFormulaTranslator } from '../../src/formulas/openFormulaTranslator.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { roundTrip } from '../serialization/helpers.ts';
import { bool, errorToken, num, text } from './formulaHarness.ts';

const write = (formula: string): string => FormulaWriter.write(Parser.parse(formula));

/** A1:A3 holds 10, 20, 30. */
function column(sheet: Worksheet): void {
  for (let r = 0; r < 3; r++) sheet.setValue(r, 0, CellValue.number((r + 1) * 10));
}

describe('inline arrays', () => {
  it.each([
    ['{1,2,3}', '{1,2,3}'],
    ['{1;2;3}', '{1;2;3}'],
    ['{ 1 , 2 ; 3 , 4 }', '{1,2;3,4}'],
    ['{-1,+2.5,"a""b";TRUE,FALSE,#N/A}', '{-1,2.5,"a""b";TRUE,FALSE,#N/A}'],
    ['SUM({1,2;3,4})', 'SUM({1,2;3,4})'],
  ])('parses and writes %s as %s', (formula, expected) => {
    expect(write(formula)).toBe(expected);
  });

  it.each(['{}', '{1,2;3}', '{A1}', '{1+2}', '{1,,2}', '{SUM(1)}', '{1'])('rejects %s', (formula) => {
    expect(Parser.tryParse(formula).ok).toBe(false);
  });

  it.each([
    ['SUM({1,2;3,4})', 10],
    ['SUM({1,"x",TRUE})', 1], // as in a range, only numbers count
    ['COUNT({1,"x",2})', 2],
    ['COUNTA({1,"x",2})', 3],
    ['AVERAGE({2;4;6})', 4],
    ['MAX({-5,-2})', -2],
    ['INDEX({1,2;3,4},2,1)', 3],
    ['INDEX({10;20;30},2)', 20],
    ['INDEX({10,20,30},3)', 30],
    ['MATCH(20,{10,20,30},0)', 2],
    ['MATCH(25,{10;20;30})', 2],
    ['HLOOKUP("b",{"a","b";1,2},2,FALSE)', 2],
    ['{7,8,9}', 7], // a single value: the top-left element, as in Calc
    ['{7,8,9}+1', 8],
    ['SUMIF({1,2,3},">1")', 5],
    ['SUMIF(A1:A3,">10",{1;2;3})', 5],
  ])('%s = %d', (formula, expected) => {
    expect(num(formula, column)).toBe(expected);
  });

  it('carries text, booleans and errors', () => {
    expect(text('INDEX({"a","b"},2)')).toBe('b');
    expect(text('VLOOKUP(2,{1,"a";2,"b"},2,FALSE)')).toBe('b');
    expect(bool('INDEX({TRUE,FALSE},1)')).toBe(true);
    expect(errorToken('INDEX({1,#N/A},2)')).toBe('#N/A');
    expect(errorToken('SUM({1,#DIV/0!})')).toBe('#DIV/0!');
    expect(errorToken('INDEX({1,2},3)')).toBe('#REF!');
  });

  it('round-trips through OpenFormula with its column (;) and row (|) separators', () => {
    expect(OpenFormulaTranslator.toOpenFormula('SUM({1,2;3,4})')).toBe('of:=SUM({1;2|3;4})');
    expect(OpenFormulaTranslator.toOpenFormula('{"a;b",TRUE;-1,#N/A}')).toBe('of:={"a;b";TRUE|-1;#N/A}');
    expect(OpenFormulaTranslator.fromOpenFormula('of:=SUM({1;2|3;4})')).toBe('SUM({1,2;3,4})');
    expect(OpenFormulaTranslator.fromOpenFormula('of:={"x|y";TRUE()|FALSE;2}')).toBe('{"x|y",TRUE;FALSE,2}');
    expect(OpenFormulaTranslator.fromOpenFormula('of:=MATCH(2;{1;2;3};0)')).toBe('MATCH(2, {1,2,3}, 0)');
  });

  it('round-trips through an .ods file', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.setFormula(0, 0, 'SUM({1,2;3,4})');
    workbook.activeSheet.setFormula(0, 1, 'INDEX({"a";"b"},2)');
    const reread = roundTrip(workbook).activeSheet;
    expect(reread.find(0, 0)?.formula).toBe('SUM({1,2;3,4})');
    expect(reread.find(0, 1)?.formula).toBe('INDEX({"a";"b"}, 2)');
  });

  it('commas inside an array do not move the argument hint', () => {
    expect(FormulaSignatureScanner.find('=MATCH(2,{1,2,3},', 17)).toEqual({
      functionName: 'MATCH',
      argumentIndex: 2,
    });
  });
});
