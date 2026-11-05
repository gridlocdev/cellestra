import { describe, expect, it } from 'vitest';
import type { FormulaCallSite } from '../../src/formulas/functions/formulaSignatureScanner.ts';
import { FormulaSignatureScanner } from '../../src/formulas/functions/formulaSignatureScanner.ts';

/** Marks the caret with '|'. */
function find(withCaret: string): FormulaCallSite | null {
  const caret = withCaret.indexOf('|');
  return FormulaSignatureScanner.find(withCaret.replace('|', ''), caret);
}

describe('FormulaSignatureScanner', () => {
  it('caret just inside the paren is argument zero', () => {
    expect(find('=SUM(|')).toEqual({ functionName: 'SUM', argumentIndex: 0 });
  });

  it('commas advance the argument index', () => {
    expect(find('=SUM(A1, B2, |')).toEqual({ functionName: 'SUM', argumentIndex: 2 });
  });

  it('nested calls report the innermost', () => {
    expect(find('=SUM(A1, AVERAGE(B1, |')).toEqual({ functionName: 'AVERAGE', argumentIndex: 1 });
  });

  it('after a closed nested call reports the outer', () => {
    expect(find('=SUM(A1, AVERAGE(B1, B2), |')).toEqual({ functionName: 'SUM', argumentIndex: 2 });
  });

  it('commas inside string literals are ignored', () => {
    expect(find('=IF(A1="a,b,c", |')).toEqual({ functionName: 'IF', argumentIndex: 1 });
  });

  it('outside any call returns null', () => {
    expect(find('=A1+|')).toBeNull();
    expect(find('=SUM(A1)|')).toBeNull(); // caret after the closing paren
  });

  it('grouping parens fall through to the enclosing call', () => {
    expect(find('=SUM((A1+B1)|')).toEqual({ functionName: 'SUM', argumentIndex: 0 });
  });
});
