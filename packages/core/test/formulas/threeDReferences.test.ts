import { describe, expect, it } from 'vitest';
import { OpenFormulaTranslator } from '../../src/formulas/openFormulaTranslator.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';

describe('Ranges across several sheets', () => {
  it.each([
    'SUM(Sheet1.A1:Sheet3.B2)',
    "SUM('My Sheet'.A1:Sheet2.B2)",
    'SUM(A1:Sheet2.B2)',
    'SUM(Sheet1.A:Sheet3.C)',
  ])('refuses %s instead of reading it as one sheet', (formula) => {
    expect(Parser.tryParse(formula).ok).toBe(false);
  });

  it('still reads a range qualified with the same sheet on both corners', () => {
    expect(Parser.tryParse('SUM(Sheet1.A1:Sheet1.B2)').ok).toBe(true);
    expect(Parser.tryParse('SUM(sheet1.A1:Sheet1.B2)').ok).toBe(true);
  });

  it('saves a loaded 3D range back unchanged', () => {
    const original = 'of:=SUM([$Sheet1.A1:$Sheet3.B2])';
    const { formula, source } = OpenFormulaTranslator.read(original);
    expect(OpenFormulaTranslator.toOpenFormula(formula, source)).toBe(original);
  });
});
