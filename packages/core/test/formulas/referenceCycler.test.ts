import { describe, expect, it } from 'vitest';
import { ReferenceCycler } from '../../src/formulas/references/referenceCycler.ts';

describe('ReferenceCycler', () => {
  it.each([
    ['=A1', '=$A$1'],
    ['=$A$1', '=A$1'],
    ['=A$1', '=$A1'],
    ['=$A1', '=A1'],
  ])('cycles %s to %s', (text, expected) => {
    // Caret at the end (just after typing the reference).
    expect(ReferenceCycler.cycle(text, text.length).text).toBe(expected);
  });

  it('cycles the reference under the caret within a larger formula', () => {
    // =A1+B2 with the caret right after B2.
    const { text, caret } = ReferenceCycler.cycle('=A1+B2', 6);
    expect(text).toBe('=A1+$B$2');
    expect(caret).toBe(8); // caret follows the grown reference
  });

  it('leaves function names alone', () => {
    expect(ReferenceCycler.cycle('=SUM(', 4).text).toBe('=SUM('); // caret inside "SUM"
  });

  it('multi-letter columns and rows cycle', () => {
    expect(ReferenceCycler.cycle('=AB12', 5).text).toBe('=$AB$12');
  });
});
