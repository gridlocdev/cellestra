import { describe, expect, it } from 'vitest';
import { FormulaWriter } from '../../src/formulas/formulaWriter.ts';
import { OpenFormulaTranslator } from '../../src/formulas/openFormulaTranslator.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { roundTrip } from '../serialization/helpers.ts';
import { errorToken, num } from './formulaHarness.ts';

/** A1:B2 = 1, 2 / 3, 4 and D1:E2 = 10, 20 / 30, 40. */
function setup(sheet: Worksheet): void {
  sheet.setValue(0, 0, CellValue.number(1));
  sheet.setValue(0, 1, CellValue.number(2));
  sheet.setValue(1, 0, CellValue.number(3));
  sheet.setValue(1, 1, CellValue.number(4));
  sheet.setValue(0, 3, CellValue.number(10));
  sheet.setValue(0, 4, CellValue.number(20));
  sheet.setValue(1, 3, CellValue.number(30));
  sheet.setValue(1, 4, CellValue.number(40));
}

describe('INDEX with an area number', () => {
  it.each([
    ['INDEX((A1:B2,D1:E2),2,2,2)', 'INDEX(A1:B2~D1:E2, 2, 2, 2)'],
    ['INDEX((A1:B2;D1:E2),1,1,1)', 'INDEX(A1:B2~D1:E2, 1, 1, 1)'],
    ['INDEX(A1:B2~D1:E2,1,2,2)', 'INDEX(A1:B2~D1:E2, 1, 2, 2)'],
  ])('parses and writes %s as %s', (formula, expected) => {
    expect(FormulaWriter.write(Parser.parse(formula))).toBe(expected);
  });

  it.each([
    ['INDEX((A1:B2,D1:E2),2,2,2)', 40],
    ['INDEX((A1:B2,D1:E2),1,2,2)', 20],
    ['INDEX((A1:B2,D1:E2),2,1,1)', 3],
    ['INDEX((A1:B2,D1:E2),2,1)', 3], // the area defaults to the first
    ['INDEX(A1:B2,2,2,1)', 4],
    ['SUM(INDEX((A1:B2,D1:E2),0,0,2))', 100],
    ['SUM(INDEX((A1:B2,D1:E2),0,2,2))', 60],
    ['SUM((INDEX((A1:B2,D1:E2),1,1,2),A1))', 11], // INDEX returns a reference
    ['INDEX({1,2;3,4},2,1,1)', 3],
  ])('%s = %d', (formula, expected) => {
    expect(num(formula, setup)).toBe(expected);
  });

  it.each([
    ['INDEX((A1:B2,D1:E2),1,1,3)', '#REF!'],
    ['INDEX((A1:B2,D1:E2),1,1,0)', '#VALUE!'], // as LibreOffice (Err:502)
    ['INDEX(A1:B2,1,1,0)', '#VALUE!'],
    ['INDEX(A1:B2,1,1,2)', '#REF!'],
    ['INDEX((A1:B2,D1:E2),3,1,2)', '#REF!'],
    ['INDEX((A1:B2,D1:E2),1,1,"x")', '#VALUE!'],
  ])('%s is %s', (formula, token) => {
    expect(errorToken(formula, setup)).toBe(token);
  });

  it('round-trips through OpenFormula', () => {
    expect(OpenFormulaTranslator.toOpenFormula('INDEX((A1:B2,D1:E2),2,2,2)')).toBe(
      'of:=INDEX([.A1:.B2]~[.D1:.E2];2;2;2)',
    );
    expect(OpenFormulaTranslator.fromOpenFormula('of:=INDEX([.A1:.B2]~[.D1:.E2];2;2;2)')).toBe(
      'INDEX(A1:B2~D1:E2, 2, 2, 2)',
    );
    expect(OpenFormulaTranslator.fromOpenFormula('of:=INDEX(([.A1:.B2];[.D1:.E2]);2;2;2)')).toBe(
      'INDEX(A1:B2~D1:E2, 2, 2, 2)',
    );
  });

  it('round-trips through an .ods file', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.setFormula(5, 0, 'INDEX(A1:B2~D1:E2, 2, 2, 2)');
    expect(roundTrip(workbook).activeSheet.find(5, 0)?.formula).toBe('INDEX(A1:B2~D1:E2, 2, 2, 2)');
  });
});
