import { describe, expect, it } from 'vitest';
import { OpenFormulaTranslator } from '../../src/formulas/openFormulaTranslator.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';
import { CellAddress } from '../../src/model/cellAddress.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { evaluate } from './formulaHarness.ts';

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

describe('OpenFormulaTranslator round-trips what it cannot rewrite', () => {
  it('writes an unparsed formula back exactly as it was read', () => {
    const original = 'of:=MYSTERY([.A1] ~ [.B1]; {1;2})';
    const { formula, source } = OpenFormulaTranslator.read(original);
    expect(OpenFormulaTranslator.toOpenFormula(formula, source)).toBe(original);
  });

  it('parses empty arguments', () => {
    const internal = OpenFormulaTranslator.fromOpenFormula('of:=IF([.A1]>0;;1)');
    expect(Parser.tryParse(internal).ok).toBe(true);
    expect(OpenFormulaTranslator.toOpenFormula(internal)).toBe('of:=IF([.A1]>0;;1)');
  });

  it.each([
    ['of:=SUM([.A:.A])', 'SUM(A:A)'],
    ['of:=SUM([.$B:.$C])', 'SUM($B:$C)'],
    ['of:=SUM([Sheet2.A:.B])', 'SUM(Sheet2.A:B)'],
    ['of:=SUM([.1:.1])', 'SUM(1:1)'],
    ['of:=SUM([.$2:.$4])', 'SUM($2:$4)'],
  ])('parses whole-column and whole-row ranges: %s', (openFormula, internal) => {
    expect(OpenFormulaTranslator.fromOpenFormula(openFormula)).toBe(internal);
    expect(OpenFormulaTranslator.toOpenFormula(internal)).toBe(openFormula);
  });
});

describe('unary minus and ^ mean the same in both directions', () => {
  it.each([
    ['(-2)^2', 4],
    ['-2^2', 4],
    ['-(2^2)', -4],
    ['2^-1', 0.5],
    ['-2^-2', 0.25],
    ['2*-3^2', 18],
  ])('%s keeps its value through OpenFormula', (formula, value) => {
    const back = OpenFormulaTranslator.fromOpenFormula(OpenFormulaTranslator.toOpenFormula(formula));
    expect(evaluate(back)).toEqual(CellValue.number(value));
    expect(evaluate(formula)).toEqual(CellValue.number(value));
  });

  it('writes the ambiguous forms with parentheses', () => {
    expect(OpenFormulaTranslator.toOpenFormula('(-2)^2')).toBe('of:=(-2)^2');
    expect(OpenFormulaTranslator.toOpenFormula('-2^2')).toBe('of:=(-2)^2');
    expect(OpenFormulaTranslator.toOpenFormula('-(2^2)')).toBe('of:=-(2^2)');
  });

  // Calc evaluates of:=-[.B2]^2 as (-B2)^2 (B2=3 gives 9): the import must keep that value.
  it("an imported signed power keeps Calc's meaning", () => {
    expect(OpenFormulaTranslator.fromOpenFormula('of:=-[.B2]^2')).toBe('-B2^2');
    expect(OpenFormulaTranslator.fromOpenFormula('of:=-([.B2]^2)')).toBe('-(B2^2)');
    expect(OpenFormulaTranslator.fromOpenFormula('of:=EXP(-[.B2]^2/2)')).toBe('EXP(-B2^2/2)');
    const back = OpenFormulaTranslator.fromOpenFormula('of:=-[.B2]^2');
    expect(OpenFormulaTranslator.toOpenFormula(back)).toBe('of:=(-[.B2])^2');
  });
});

describe('sheet names that need quoting', () => {
  it.each([
    ['2024', "'2024'.A1", "of:=['2024'.A1]"],
    ['AB12', "'AB12'.A1", "of:=['AB12'.A1]"],
    ['TRUE', "'TRUE'.A1", "of:=['TRUE'.A1]"],
    ['Sheet1', 'Sheet1.A1', 'of:=[Sheet1.A1]'],
  ])('%s', (sheet, a1, openFormula) => {
    const written = new CellAddress(0, 0, false, false, sheet).toString();
    expect(written).toBe(a1);
    expect(Parser.tryParse(written).ok).toBe(true);
    expect(OpenFormulaTranslator.toOpenFormula(written)).toBe(openFormula);
    expect(OpenFormulaTranslator.fromOpenFormula(openFormula)).toBe(a1);
  });
});

describe('Excel 2010+ functions in the COM.MICROSOFT namespace', () => {
  it('CONCAT is written with the prefix LibreOffice needs and read back without it', () => {
    expect(OpenFormulaTranslator.toOpenFormula('CONCAT("a",A1)')).toBe('of:=COM.MICROSOFT.CONCAT("a";[.A1])');
    expect(OpenFormulaTranslator.fromOpenFormula('of:=COM.MICROSOFT.CONCAT("a";[.A1])')).toBe(
      'CONCAT("a", A1)',
    );
    expect(OpenFormulaTranslator.fromOpenFormula('of:=com.microsoft.concat("a")')).toBe('CONCAT("a")');
    expect(evaluate(OpenFormulaTranslator.fromOpenFormula('of:=COM.MICROSOFT.CONCAT("a";"b")'))).toEqual(
      CellValue.text('ab'),
    );
  });

  it.each(['TEXTJOIN', 'IFS', 'SWITCH', 'MAXIFS', 'MINIFS', 'XLOOKUP'])('%s round-trips', (name) => {
    const openFormula = OpenFormulaTranslator.toOpenFormula(`${name}(1,2)`);
    expect(openFormula).toBe(`of:=COM.MICROSOFT.${name}(1;2)`);
    expect(OpenFormulaTranslator.fromOpenFormula(openFormula)).toBe(`${name}(1, 2)`);
  });

  it('leaves the prefix alone inside strings and on other functions', () => {
    expect(OpenFormulaTranslator.fromOpenFormula('of:="COM.MICROSOFT.CONCAT("')).toBe(
      '"COM.MICROSOFT.CONCAT("',
    );
    expect(OpenFormulaTranslator.toOpenFormula('SUM(1)')).toBe('of:=SUM(1)');
  });
});
