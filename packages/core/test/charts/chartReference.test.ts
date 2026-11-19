import { describe, expect, it } from 'vitest';
import { expandWholeLines, parseChartRange } from '../../src/index.ts';

describe('expandWholeLines', () => {
  it('writes whole columns out as cells', () => {
    expect(expandWholeLines('Data.A:C')).toBe('Data.A1:C1048576');
    expect(expandWholeLines(' A:A ')).toBe('A1:A1048576');
    expect(expandWholeLines('Data.$B:$B')).toBe('Data.$B$1:$B$1048576');
    expect(expandWholeLines("'My Data'.A:A")).toBe("'My Data'.A1:A1048576");
  });

  it('writes whole rows out as cells', () => {
    expect(expandWholeLines('Data.2:3')).toBe('Data.A2:XFD3');
    expect(expandWholeLines('$1:$1')).toBe('$A$1:$XFD$1');
  });

  it('leaves cell ranges and nonsense as typed, trimmed', () => {
    expect(expandWholeLines(' Data.A1:B2 ')).toBe('Data.A1:B2');
    expect(expandWholeLines('Data.A1')).toBe('Data.A1');
    expect(expandWholeLines('A:2')).toBe('A:2');
    expect(expandWholeLines('0:3')).toBe('0:3');
    expect(expandWholeLines('ZZZZ:A')).toBe('ZZZZ:A');
  });
});

describe('parseChartRange', () => {
  it('parses whole columns and rows as well as cell ranges', () => {
    const column = parseChartRange('Data.B:B')!;
    expect([column.sheet, column.left, column.right, column.top, column.rowCount]).toEqual([
      'Data',
      1,
      1,
      0,
      1048576,
    ]);
    expect(parseChartRange('Data.2:2')!.columnCount).toBe(16384);
    expect(parseChartRange('A1:A3')!.rowCount).toBe(3);
  });

  it('is null for empty text and text that names no range', () => {
    expect(parseChartRange('')).toBeNull();
    expect(parseChartRange(null)).toBeNull();
    expect(parseChartRange('Sales')).toBeNull();
  });
});
