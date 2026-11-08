import { describe, expect, it } from 'vitest';
import { Expr } from '../../src/formulas/ast/expr.ts';
import { FormulaWriter } from '../../src/formulas/formulaWriter.ts';
import { Lexer } from '../../src/formulas/lexing/lexer.ts';
import { OpenFormulaTranslator } from '../../src/formulas/openFormulaTranslator.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';
import { CellAddress } from '../../src/model/cellAddress.ts';
import { bool, errorToken } from './formulaHarness.ts';

describe('error literals', () => {
  it.each(['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!'])(
    '%s lexes and evaluates',
    (token) => {
      expect(Lexer.tokenize(`1+${token.toLowerCase()}`).map((t) => t.type)).toEqual([
        'number',
        'plus',
        'error',
        'endOfInput',
      ]);
      expect(errorToken(token)).toBe(token);
      expect(errorToken(`1+${token}`)).toBe(token);
      expect(bool(`IFERROR(${token},TRUE)`)).toBe(true);
    },
  );

  it('#REF! parses to the broken-reference node, so a written formula reads back', () => {
    const broken = Expr.binary('add', Expr.reference(CellAddress.parse('A1')), Expr.refError);
    const written = FormulaWriter.write(broken);
    expect(written).toBe('A1+#REF!');
    expect(Parser.parse(written)).toEqual(broken);
  });

  it('an unknown # word is still a parse error', () => {
    expect(Parser.tryParse('#FOO!').ok).toBe(false);
  });

  it.each(['of:=[.A1]+[#REF!]', 'of:=[.A1]+[.#REF!]', 'of:=[.A1]+[.A1:.#REF!]', 'of:=[.A1]+[$#REF!.A1]'])(
    'imports the broken reference in %s',
    (openFormula) => {
      expect(OpenFormulaTranslator.fromOpenFormula(openFormula)).toBe('A1+#REF!');
    },
  );

  it('exports a broken reference and error literals as LibreOffice reads them', () => {
    expect(OpenFormulaTranslator.toOpenFormula('A1+#REF!')).toBe('of:=[.A1]+[#REF!]');
    expect(OpenFormulaTranslator.toOpenFormula('IFERROR(#N/A,1)')).toBe('of:=IFERROR(#N/A;1)');
  });
});
