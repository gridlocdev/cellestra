import { describe, expect, it } from 'vitest';
import { CellRef } from '../../src/dependencies/cellRef.ts';
import { RecalculationEngine } from '../../src/dependencies/recalculationEngine.ts';
import type { Expr } from '../../src/formulas/ast/expr.ts';
import { FormulaWriter } from '../../src/formulas/formulaWriter.ts';
import { OpenFormulaTranslator } from '../../src/formulas/openFormulaTranslator.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';
import { ReferenceAdjuster } from '../../src/formulas/references/referenceAdjuster.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { roundTrip } from '../serialization/helpers.ts';
import { errorToken, num } from './formulaHarness.ts';

const write = (expr: Expr): string => FormulaWriter.write(expr);
const parse = (formula: string): Expr => Parser.parse(formula);

/** A1:C3 holds 1..9 row by row. */
function grid(sheet: Worksheet): void {
  for (let r = 0; r < 3; r++)
    for (let c = 0; c < 3; c++) sheet.setValue(r, c, CellValue.number(r * 3 + c + 1));
}

describe('reference union (~) and intersection (!)', () => {
  it.each([
    ['SUM(A1:A2~C1)', 'SUM(A1:A2~C1)'],
    ['SUM((A1:A2,C1))', 'SUM(A1:A2~C1)'], // a list in bare parentheses is a union, as in Calc
    ['SUM((A1:A2;C1))', 'SUM(A1:A2~C1)'],
    ['SUM(A1:C3!B2:D4)', 'SUM(A1:C3!B2:D4)'],
    ['SUM(A1:B2~B2:C3!C3)', 'SUM(A1:B2~B2:C3!C3)'], // ! binds tighter than ~
    ['SUM((A1~B1)!A1:C1)', 'SUM((A1~B1)!A1:C1)'],
    ['-A1~B1', '-A1~B1'], // reference operators bind tighter than a sign
  ])('parses and writes %s as %s', (formula, expected) => {
    expect(write(parse(formula))).toBe(expected);
  });

  it('rejects a list of values in bare parentheses', () => {
    expect(Parser.tryParse('SUM((1,2))').ok).toBe(false);
  });

  it.each([
    ['SUM(A1:A2~C1)', 1 + 4 + 3],
    ['SUM(A1~A1)', 2], // Calc counts an area listed twice twice
    ['COUNT(A1:A3~B1:B3)', 6],
    ['AVERAGE(A1~C3)', 5],
    ['MAX(A1~B1~C1)', 3],
    ['SUM((A1:A2,C1))', 8],
    ['SUM(A1:C3!B2:D4)', 5 + 6 + 8 + 9],
    ['SUM(A1:B2~B2:C3!C3)', 1 + 2 + 4 + 5 + 9],
    ['A1:C1!B1:B3', 2], // one cell left: a single value
    ['COUNTA(A1:A3!A2:C2)', 1],
  ])('%s = %d', (formula, expected) => {
    expect(num(formula, grid)).toBe(expected);
  });

  it('ranges that do not meet intersect to #NULL!', () => {
    expect(errorToken('SUM(A1:A3!C1:C3)', grid)).toBe('#NULL!');
  });

  it('a union where one value is needed is #VALUE!', () => {
    expect(errorToken('A1~B1', grid)).toBe('#VALUE!');
  });

  it('a union of a value is #VALUE!', () => {
    expect(errorToken('SUM(A1~(1+2))', grid)).toBe('#VALUE!');
  });

  it('round-trips through OpenFormula', () => {
    expect(OpenFormulaTranslator.toOpenFormula('SUM(A1:A2~C1)')).toBe('of:=SUM([.A1:.A2]~[.C1])');
    expect(OpenFormulaTranslator.toOpenFormula('SUM(A1:C3!B2:D4)')).toBe('of:=SUM([.A1:.C3]![.B2:.D4])');
    expect(OpenFormulaTranslator.fromOpenFormula('of:=SUM([.A1:.B2]~[.C3])')).toBe('SUM(A1:B2~C3)');
    expect(OpenFormulaTranslator.fromOpenFormula('of:=SUM([.A1:.C3]![.B2:.D4])')).toBe('SUM(A1:C3!B2:D4)');
    expect(OpenFormulaTranslator.fromOpenFormula('of:=SUM(([.A1];[.B1]))')).toBe('SUM(A1~B1)');
  });

  it('round-trips through an .ods file', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setFormula(5, 0, 'SUM(A1:A2~C1)');
    sheet.setFormula(5, 1, 'SUM(A1:C3!B2:D4)');
    const reread = roundTrip(workbook).activeSheet;
    expect(reread.find(5, 0)?.formula).toBe('SUM(A1:A2~C1)');
    expect(reread.find(5, 1)?.formula).toBe('SUM(A1:C3!B2:D4)');
  });

  it('tracks every area as a dependency', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    grid(sheet);
    sheet.setFormula(5, 0, 'SUM(A1~C3)');
    sheet.setFormula(5, 1, 'SUM(A1:C3!C2:D4)');
    const engine = new RecalculationEngine(workbook);
    engine.rebuildGraph();
    engine.recalculateAll();
    expect(CellValue.asNumber(sheet.getValue(5, 0))).toBe(10);
    expect(CellValue.asNumber(sheet.getValue(5, 1))).toBe(15);

    sheet.setValue(2, 2, CellValue.number(100));
    engine.recordEdits([CellRef.of(sheet.name, 2, 2)]);
    expect(CellValue.asNumber(sheet.getValue(5, 0))).toBe(101);
    expect(CellValue.asNumber(sheet.getValue(5, 1))).toBe(106);
  });

  it('moves every area on insert and breaks only the deleted one', () => {
    expect(write(ReferenceAdjuster.forInsertRows(parse('SUM(A1:A2~C3)'), 'Sheet1', 'Sheet1', 1, 1))).toBe(
      'SUM(A1:A3~C4)',
    );
    expect(write(ReferenceAdjuster.forDeleteColumns(parse('SUM(A1:A2~C1)'), 'Sheet1', 'Sheet1', 2, 1))).toBe(
      'SUM(A1:A2~#REF!)',
    );
    expect(
      write(ReferenceAdjuster.forInsertColumns(parse('SUM(A1:C3!B2:D4)'), 'Sheet1', 'Sheet1', 0, 1)),
    ).toBe('SUM(B1:D3!C2:E4)');
    expect(write(ReferenceAdjuster.forCopy(parse('SUM(A1~$B$1)'), 0, 1))).toBe('SUM(A2~$B$1)');
  });
});
