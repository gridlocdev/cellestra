import { describe, expect, it } from 'vitest';
import { OpenFormulaTranslator } from '../../src/formulas/openFormulaTranslator.ts';

describe('OpenFormulaTranslator', () => {
  it.each([
    ['SUM(A1:A3)', 'of:=SUM([.A1:.A3])'],
    ['A1+B2', 'of:=[.A1]+[.B2]'],
    ['IF(A1>0, B1, C1)', 'of:=IF([.A1]>0;[.B1];[.C1])'],
    ['$A$1*2', 'of:=[.$A$1]*2'],
    ['Sheet2.A1', 'of:=[Sheet2.A1]'],
    ['"a;b"', 'of:="a;b"'],
  ])('toOpenFormula(%s) = %s', (internalFormula, expected) => {
    expect(OpenFormulaTranslator.toOpenFormula(internalFormula)).toBe(expected);
  });

  it.each([
    ['of:=SUM([.A1:.A3])', 'SUM(A1:A3)'],
    ['of:=[.A1]+[.B2]', 'A1+B2'],
    ['of:=IF([.A1]>0;[.B1];[.C1])', 'IF(A1>0, B1, C1)'],
    ['of:=[.$A$1]*2', '$A$1*2'],
    ["of:=['My Sheet'.A1]", "'My Sheet'.A1"],
  ])('fromOpenFormula(%s) = %s', (openFormula, expected) => {
    expect(OpenFormulaTranslator.fromOpenFormula(openFormula)).toBe(expected);
  });

  it.each(['SUM(A1:A3)', 'IF(A1>0, A1*2, 0)', '$A$1+B$2-C3', 'Sheet2.A1+A1', '"x;y"&A1'])(
    'round-trips %s through OpenFormula',
    (internalFormula) => {
      const odf = OpenFormulaTranslator.toOpenFormula(internalFormula);
      expect(OpenFormulaTranslator.fromOpenFormula(odf)).toBe(internalFormula);
    },
  );
});
