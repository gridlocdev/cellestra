import { describe, expect, it } from 'vitest';
import { RecalculationEngine } from '../../src/dependencies/recalculationEngine.ts';
import { OpenFormulaTranslator } from '../../src/formulas/openFormulaTranslator.ts';
import {
  MAX_FORMULA_DEPTH,
  MAX_FORMULA_NESTING,
  MAX_FORMULA_TOKENS,
  Parser,
} from '../../src/formulas/parsing/parser.ts';
import { FormulaParseException } from '../../src/formulas/parsing/parseError.ts';
import { Workbook } from '../../src/model/workbook.ts';

const parens = (n: number): string => '('.repeat(n) + '1' + ')'.repeat(n);
const ifs = (n: number): string => 'IF(1,'.repeat(n) + '1' + ')'.repeat(n);
const signs = (n: number): string => '-'.repeat(n) + '1';
const chain = (n: number): string => Array<string>(n).fill('1').join('+');

describe('formula size limits', () => {
  it.each([
    ['parentheses', parens(MAX_FORMULA_NESTING)],
    ['nested IFs', ifs(MAX_FORMULA_NESTING)],
    ['signs', signs(MAX_FORMULA_NESTING)],
    ['an operator chain', chain(MAX_FORMULA_DEPTH)],
  ])('%s up to the limit parse, write and evaluate', (_, formula) => {
    expect(Parser.tryParse(formula).ok).toBe(true);
    expect(OpenFormulaTranslator.fromOpenFormula(OpenFormulaTranslator.toOpenFormula(formula))).toBeTruthy();
    const workbook = Workbook.createNew();
    workbook.activeSheet.setFormula(0, 0, formula);
    const engine = new RecalculationEngine(workbook);
    engine.rebuildGraph();
    engine.recalculateAll();
    expect(workbook.activeSheet.getValue(0, 0).kind).toBe('number');
  });

  it.each([
    ['1000 parentheses', parens(1000)],
    ['IF nested 2000 deep', ifs(2000)],
    ['50000 signs', signs(50_000)],
    ['a 2000-term chain', chain(2000)],
    ['100000 terms', chain(100_000)],
    ['parentheses past the limit', parens(MAX_FORMULA_NESTING + 1)],
  ])('%s is a parse error, not a crash', (_, formula) => {
    expect(() => Parser.parse(formula)).toThrow(FormulaParseException);
    const parsed = Parser.tryParse(formula);
    expect(parsed.ok).toBe(false);
    // Everything downstream of the parser degrades to an error value or the text as typed.
    expect(OpenFormulaTranslator.toOpenFormula(formula)).toBe('of:=' + formula);
    const workbook = Workbook.createNew();
    workbook.activeSheet.setFormula(0, 0, formula);
    const engine = new RecalculationEngine(workbook);
    engine.rebuildGraph();
    engine.recalculateAll();
    expect(workbook.activeSheet.getValue(0, 0).kind).toBe('error');
  });

  it('caps the token count like Calc', () => {
    const tooLong = Parser.tryParse(chain(MAX_FORMULA_TOKENS / 2 + 1));
    expect(tooLong.ok).toBe(false);
    if (!tooLong.ok) expect(tooLong.error.reason).toMatch(/too long/);
  });
});
