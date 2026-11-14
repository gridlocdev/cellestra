import { describe, expect, it } from 'vitest';
import { FormulaBarDraft } from '../src/components/formulaBarDraft.ts';

describe('FormulaBarDraft', () => {
  it('follows the active cell until the user types, so focus alone commits nothing', () => {
    const d = new FormulaBarDraft('old');
    d.sync('new sheet value', 'Sheet2|0|0'); // focused, then Ctrl+PageDown
    expect(d.text).toBe('new sheet value');
    expect(d.take('new sheet value')).toBeNull();
  });

  it('commits what the user typed', () => {
    const d = new FormulaBarDraft('1');
    d.input('12', 'Sheet1|0|0');
    d.sync('1', 'Sheet1|0|0'); // a recalculation elsewhere doesn't wipe the typing
    expect(d.take('1')).toBe('12');
    expect(d.dirty).toBe(false);
  });

  it('drops typing when the cell it began in is no longer the active one', () => {
    const d = new FormulaBarDraft('a');
    d.input('typed', 'Sheet1|0|0');
    d.sync('b', 'Sheet2|0|0');
    expect(d.text).toBe('b');
    expect(d.take('b')).toBeNull();
  });

  it('commits nothing when the text ends up unchanged', () => {
    const d = new FormulaBarDraft('same');
    d.input('same', 'Sheet1|0|0');
    expect(d.take('same')).toBeNull();
  });
});
