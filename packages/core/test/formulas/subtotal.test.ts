import { describe, expect, it } from 'vitest';
import { RecalculationEngine } from '../../src/dependencies/recalculationEngine.ts';
import { Evaluator } from '../../src/formulas/evaluation/evaluator.ts';
import { WorkbookEvaluationContext } from '../../src/formulas/evaluation/workbookEvaluationContext.ts';
import { FunctionRegistry } from '../../src/formulas/functions/functionRegistry.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const registry = FunctionRegistry.createDefault();

/** A1:A4 hold 10, 20, 30, 40; A5 holds a subtotal of them; B1 holds text. */
function workbook(): Workbook {
  const wb = Workbook.createNew();
  const sheet = wb.worksheets[0]!;
  [10, 20, 30, 40].forEach((n, row) => sheet.setValue(row, 0, CellValue.number(n)));
  sheet.setFormula(4, 0, 'SUBTOTAL(9, A1:A4)', CellValue.number(100));
  sheet.setValue(0, 1, CellValue.text('label'));
  return wb;
}

function evaluate(formula: string, wb = workbook()): CellValue {
  const context = new WorkbookEvaluationContext(wb, wb.worksheets[0]!.name, 'en-US', undefined, {
    row: 9,
    column: 9,
  });
  return new Evaluator(context, registry).evaluate(Parser.parse(formula));
}

const number = (value: CellValue): number => (value.kind === 'number' ? value.value : NaN);

describe('SUBTOTAL', () => {
  it.each([
    ['SUBTOTAL(1, A1:A4)', 25],
    ['SUBTOTAL(2, A1:B4)', 4],
    ['SUBTOTAL(3, A1:B4)', 5],
    ['SUBTOTAL(4, A1:A4)', 40],
    ['SUBTOTAL(5, A1:A4)', 10],
    ['SUBTOTAL(6, A1:A4)', 240000],
    ['SUBTOTAL(9, A1:A4)', 100],
    ['SUBTOTAL(9, A1:A2, A4)', 70],
    ['SUBTOTAL(10, A1:A4)', 500 / 3],
    ['SUBTOTAL(11, A1:A4)', 125],
    ['SUBTOTAL(109, A1:A4)', 100],
  ])('%s is %d', (formula, expected) => {
    expect(number(evaluate(formula))).toBeCloseTo(expected, 9);
  });

  it('works out the standard deviations', () => {
    expect(number(evaluate('SUBTOTAL(7, A1:A4)'))).toBeCloseTo(Math.sqrt(500 / 3), 9);
    expect(number(evaluate('SUBTOTAL(8, A1:A4)'))).toBeCloseTo(Math.sqrt(125), 9);
  });

  it('leaves out other subtotals inside its range', () => {
    expect(number(evaluate('SUBTOTAL(9, A1:A5)'))).toBe(100);
    expect(number(evaluate('SUM(A1:A5)'))).toBe(200);
  });

  it('counts hidden rows for 1–11 and leaves them out for 101–111', () => {
    const wb = workbook();
    wb.worksheets[0]!.getOrCreateRow(1).hidden = true;
    expect(number(evaluate('SUBTOTAL(9, A1:A4)', wb))).toBe(100);
    expect(number(evaluate('SUBTOTAL(109, A1:A4)', wb))).toBe(80);
    expect(number(evaluate('SUBTOTAL(102, A1:A4)', wb))).toBe(3);
  });

  it('always leaves out rows a filter hid', () => {
    const wb = workbook();
    const sheet = wb.worksheets[0]!;
    sheet.activeFilter = CellRange.parse('A1:A4');
    sheet.getOrCreateRow(2).hidden = true;
    expect(number(evaluate('SUBTOTAL(9, A1:A4)', wb))).toBe(70);
    expect(number(evaluate('SUBTOTAL(109, A1:A4)', wb))).toBe(70);
  });

  it('is an error for an unknown function number', () => {
    expect(evaluate('SUBTOTAL(12, A1:A4)').kind).toBe('error');
    expect(evaluate('SUBTOTAL(0, A1:A4)').kind).toBe('error');
  });

  it('recalculates when a row is hidden', () => {
    const wb = workbook();
    const sheet = wb.worksheets[0]!;
    sheet.setFormula(6, 0, 'SUBTOTAL(109, A1:A4)');
    const engine = new RecalculationEngine(wb);
    engine.rebuildGraph();
    engine.recalculateAll();
    expect(number(sheet.getValue(6, 0))).toBe(100);
    sheet.getOrCreateRow(0).hidden = true;
    engine.recalculateVolatile();
    expect(number(sheet.getValue(6, 0))).toBe(90);
  });
});
