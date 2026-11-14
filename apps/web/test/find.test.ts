import { describe, expect, it } from 'vitest';
import { CellValue, Workbook } from '@cellestra/core';
import { findMatches, nextMatch, replaceIn, replaceableText } from '../src/app/find.ts';

describe('find helpers', () => {
  const sheet = () => {
    const s = Workbook.createNew().worksheets[0]!;
    s.setValue(0, 0, CellValue.text('Total'));
    s.setValue(0, 3, CellValue.number(42));
    s.setValue(2, 1, CellValue.text('subtotal'));
    return s;
  };
  const shown = (_r: number, _c: number, cell: { value: CellValue }) =>
    cell.value.kind === 'text'
      ? cell.value.value
      : cell.value.kind === 'number'
        ? String(cell.value.value)
        : '';

  it('matches shown text without regard to case', () => {
    expect(findMatches(sheet(), 'TOTAL', shown).map((m) => m.address)).toEqual(['A1', 'B3']);
    expect(findMatches(sheet(), '42', shown).map((m) => m.address)).toEqual(['D1']);
  });

  it('steps forwards and backwards with wrap-around', () => {
    const m = findMatches(sheet(), 'total', shown);
    expect(nextMatch(m, { row: 0, column: 0 })?.address).toBe('B3');
    expect(nextMatch(m, { row: 2, column: 1 })?.address).toBe('A1');
    expect(nextMatch(m, { row: 0, column: 0 }, true)?.address).toBe('B3');
    expect(nextMatch([], { row: 0, column: 0 })).toBeNull();
  });

  it('replaces every occurrence literally, ignoring case', () => {
    expect(replaceIn('a.b A.B', 'a.b', 'x')).toBe('x x');
    expect(replaceIn('abc', 'z', 'x')).toBeNull();
    expect(replaceIn('$1', '$', '$$')).toBe('$$1');
  });

  it('edits only text and formulas', () => {
    const s = sheet();
    s.setFormula(4, 0, 'SUM(D1)', CellValue.number(42));
    expect(replaceableText(s.find(0, 0))).toBe('Total');
    expect(replaceableText(s.find(0, 3))).toBeNull();
    expect(replaceableText(s.find(4, 0))).toBe('=SUM(D1)');
    expect(replaceableText(undefined)).toBeNull();
  });
});
