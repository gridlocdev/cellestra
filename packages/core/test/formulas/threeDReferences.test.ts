import { describe, expect, it } from 'vitest';
import { CellRef } from '../../src/dependencies/cellRef.ts';
import { RecalculationEngine } from '../../src/dependencies/recalculationEngine.ts';
import { DeleteSheetCommand } from '../../src/editing/deleteSheetCommand.ts';
import { RenameSheetCommand } from '../../src/editing/renameSheetCommand.ts';
import type { Expr } from '../../src/formulas/ast/expr.ts';
import { Evaluator } from '../../src/formulas/evaluation/evaluator.ts';
import { WorkbookEvaluationContext } from '../../src/formulas/evaluation/workbookEvaluationContext.ts';
import { FormulaWriter } from '../../src/formulas/formulaWriter.ts';
import { FunctionRegistry } from '../../src/formulas/functions/functionRegistry.ts';
import { OpenFormulaTranslator } from '../../src/formulas/openFormulaTranslator.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';
import { ReferenceAdjuster } from '../../src/formulas/references/referenceAdjuster.ts';
import { Worksheet } from '../../src/model/worksheet.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { roundTrip } from '../serialization/helpers.ts';

const registry = FunctionRegistry.createDefault();
const write = (expr: Expr): string => FormulaWriter.write(expr);
const parse = (formula: string): Expr => Parser.parse(formula);

/** Sheet1..Sheet4; on sheet n, A1:B2 holds n, 10n, 100n and an empty B2. */
function workbook(): Workbook {
  const wb = Workbook.createNew();
  for (const name of ['Sheet2', 'Sheet3', 'Sheet4']) wb.addSheet(name);
  wb.worksheets.forEach((sheet, i) => {
    const n = i + 1;
    sheet.setValue(0, 0, CellValue.number(n));
    sheet.setValue(0, 1, CellValue.number(10 * n));
    sheet.setValue(1, 0, CellValue.number(100 * n));
  });
  return wb;
}

function evaluate(formula: string, wb = workbook(), sheet = 'Sheet4'): CellValue {
  const context = new WorkbookEvaluationContext(wb, sheet, 'en-US', undefined, { row: 9, column: 9 });
  return new Evaluator(context, registry).evaluate(Parser.parse(formula));
}

describe('ranges across several sheets (3D references)', () => {
  it.each([
    ['SUM(Sheet1.A1:Sheet3.B2)', 'SUM(Sheet1.A1:Sheet3.B2)'],
    ['SUM($Sheet1.A1:$Sheet3.B2)', 'SUM(Sheet1.A1:Sheet3.B2)'],
    ['SUM(Sheet1.$A$1:Sheet3.B$2)', 'SUM(Sheet1.$A$1:Sheet3.B$2)'],
    ['SUM(Sheet1.A1:Sheet3.A1)', 'SUM(Sheet1.A1:Sheet3.A1)'],
    ["SUM('My Sheet'.A1:'Other one'.B2)", "SUM('My Sheet'.A1:'Other one'.B2)"],
    ['SUM(Sheet1.A:Sheet3.C)', 'SUM(Sheet1.A:Sheet3.C)'],
    ['SUM(Sheet1.A1:Sheet1.B2)', 'SUM(Sheet1.A1:B2)'], // one sheet: an ordinary range
  ])('parses and writes %s as %s', (formula, expected) => {
    expect(write(parse(formula))).toBe(expected);
  });

  it('still refuses a range whose start names no sheet but whose end does', () => {
    expect(Parser.tryParse('SUM(A1:Sheet2.B2)').ok).toBe(false);
  });

  it.each([
    ['SUM(Sheet1.A1:Sheet3.B2)', 111 * 6],
    ['SUM(Sheet3.A1:Sheet1.B2)', 111 * 6], // sheets in either order
    ['SUM(Sheet2.A1:Sheet3.A1)', 5],
    ['COUNT(Sheet1.A1:Sheet3.B2)', 9],
    ['COUNTA(Sheet1.A1:Sheet3.B2)', 9],
    ['AVERAGE(Sheet1.A1:Sheet3.A1)', 2],
    ['MIN(Sheet1.A1:Sheet3.B2)', 1],
    ['MAX(Sheet1.A1:Sheet3.B2)', 300],
    ['SUM(Sheet1.A1:Sheet3.A1~Sheet4.A1)', 10],
  ])('%s = %d', (formula, expected) => {
    expect(evaluate(formula)).toEqual(CellValue.number(expected));
  });

  it('reads the sheets between the two by their current order', () => {
    const wb = workbook();
    wb.moveSheet(3, 1); // Sheet1, Sheet4, Sheet2, Sheet3
    expect(evaluate('SUM(Sheet1.A1:Sheet2.A1)', wb, 'Sheet3')).toEqual(CellValue.number(7));
  });

  it.each([
    ['Sheet1.A1:Sheet3.A1', '#VALUE!'], // several cells where one value is needed
    ['ABS(Sheet1.A1:Sheet3.A1)', '#VALUE!'],
    ['VLOOKUP(1,Sheet1.A1:Sheet3.B2,2,0)', '#VALUE!'],
    ['SUM(Sheet1.A1:Nowhere.B2)', '#REF!'],
    ['SUM(Nowhere.A1:Sheet1.B2)', '#REF!'],
  ])('%s is %s', (formula, token) => {
    const value = evaluate(formula);
    expect(value.kind === 'error' ? value.error.token : value.kind).toBe(token);
  });

  it('round-trips through OpenFormula as LibreOffice writes it', () => {
    const original = 'of:=SUM([$Sheet1.A1:$Sheet3.B2])';
    expect(OpenFormulaTranslator.fromOpenFormula(original)).toBe('SUM(Sheet1.A1:Sheet3.B2)');
    expect(OpenFormulaTranslator.toOpenFormula('SUM(Sheet1.A1:Sheet3.B2)')).toBe(original);
    const quoted = "of:=SUM([$'My Sheet'.$A$1:$Other.B2])";
    expect(OpenFormulaTranslator.fromOpenFormula(quoted)).toBe("SUM('My Sheet'.$A$1:Other.B2)");
    expect(OpenFormulaTranslator.toOpenFormula("SUM('My Sheet'.$A$1:Other.B2)")).toBe(quoted);
    expect(OpenFormulaTranslator.toOpenFormula('SUM(Sheet1.A:Sheet3.C)')).toBe(
      'of:=SUM([$Sheet1.A:$Sheet3.C])',
    );
  });

  it('round-trips through an .ods file', () => {
    const wb = workbook();
    wb.worksheets[3]!.setFormula(5, 0, 'SUM(Sheet1.A1:Sheet3.B2)');
    const reread = roundTrip(wb).worksheets[3]!;
    expect(reread.find(5, 0)?.formula).toBe('SUM(Sheet1.A1:Sheet3.B2)');
  });

  it('recalculates when a cell on a sheet in the middle changes', () => {
    const wb = workbook();
    const home = wb.worksheets[3]!;
    home.setFormula(5, 0, 'SUM(Sheet1.A1:Sheet3.A1)');
    const engine = new RecalculationEngine(wb);
    engine.rebuildGraph();
    engine.recalculateAll();
    expect(home.getValue(5, 0)).toEqual(CellValue.number(6));

    wb.worksheets[1]!.setValue(0, 0, CellValue.number(20));
    engine.recordEdits([CellRef.of('Sheet2', 0, 0)]);
    expect(home.getValue(5, 0)).toEqual(CellValue.number(24));

    const inserted = new Worksheet('Middle');
    inserted.setValue(0, 0, CellValue.number(1000));
    wb.insertSheet(1, inserted);
    engine.recordStructuralChange();
    expect(home.getValue(5, 0)).toEqual(CellValue.number(1024));
    inserted.setValue(0, 0, CellValue.number(2000));
    engine.recordEdits([CellRef.of('Middle', 0, 0)]);
    expect(home.getValue(5, 0)).toEqual(CellValue.number(2024));
  });

  it('renames either end with its sheet', () => {
    const wb = workbook();
    const home = wb.worksheets[3]!;
    home.setFormula(5, 0, 'SUM(Sheet1.A1:Sheet3.B2)');
    new RenameSheetCommand('Sheet1', 'First').apply(wb);
    new RenameSheetCommand('Sheet3', 'Last one').apply(wb);
    new RenameSheetCommand('Sheet2', 'Inside').apply(wb);
    expect(home.find(5, 0)?.formula).toBe("SUM(First.A1:'Last one'.B2)");
  });

  it('keeps the range when a sheet in the middle is deleted and narrows it when an end is', () => {
    const wb = workbook();
    const home = wb.worksheets[3]!;
    home.setFormula(5, 0, 'SUM(Sheet1.A1:Sheet3.B2)');
    new DeleteSheetCommand('Sheet2').apply(wb);
    expect(home.find(5, 0)?.formula).toBe('SUM(Sheet1.A1:Sheet3.B2)');
    new DeleteSheetCommand('Sheet1').apply(wb);
    expect(home.find(5, 0)?.formula).toBe('SUM(Sheet3.A1:B2)');
    new DeleteSheetCommand('Sheet3').apply(wb);
    expect(home.find(5, 0)?.formula).toBe('SUM(#REF!)');
  });

  it('moves with a copy, and stays put when rows change on one of its sheets', () => {
    expect(write(ReferenceAdjuster.forCopy(parse('SUM(Sheet1.A1:Sheet3.$B$2)'), 1, 1))).toBe(
      'SUM(Sheet1.B2:Sheet3.$B$2)',
    );
    // As in Calc: an insert on one sheet only moves a reference that lies wholly on that sheet.
    expect(
      write(ReferenceAdjuster.forInsertRows(parse('SUM(Sheet1.A1:Sheet3.B2)'), 'Sheet4', 'Sheet1', 0, 1)),
    ).toBe('SUM(Sheet1.A1:Sheet3.B2)');
  });
});
