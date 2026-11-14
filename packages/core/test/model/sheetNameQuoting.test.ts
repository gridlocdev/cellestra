import { describe, expect, it } from 'vitest';
import { Parser } from '../../src/formulas/parsing/parser.ts';
import { FormulaWriter } from '../../src/formulas/formulaWriter.ts';
import { CellAddress, quoteSheetIfNeeded } from '../../src/model/cellAddress.ts';

describe('Sheet name quoting', () => {
  it.each(['m²', 'x³', 'Area½', 'Café', 'Sheet_1', 'Données', 'Ⅻ'])(
    'writes a reference on %s that parses back to the same sheet',
    (sheet) => {
      const text = new CellAddress(0, 0, false, false, sheet).toString();
      const parsed = Parser.tryParse(text);
      expect(parsed.ok).toBe(true);
      if (parsed.ok) expect(FormulaWriter.write(parsed.expr)).toBe(text);
    },
  );

  it('quotes names with characters a plain name cannot hold', () => {
    expect(quoteSheetIfNeeded('m²')).toBe("'m²'");
    expect(quoteSheetIfNeeded('Café')).toBe('Café');
  });
});
