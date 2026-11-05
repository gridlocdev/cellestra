import { describe, expect, it } from 'vitest';
import { Axis } from '../src/grid/axis.ts';
import { HEADER_HEIGHT, HEADER_WIDTH, Viewport } from '../src/grid/viewport.ts';

const make = (over: Partial<ConstructorParameters<typeof Viewport>[0]> = {}) =>
  new Viewport({
    columns: new Axis(16384, 64, []),
    rows: new Axis(1048576, 20, []),
    frozenColumns: 0,
    frozenRows: 0,
    scrollX: 0,
    scrollY: 0,
    zoom: 1,
    width: HEADER_WIDTH + 64 * 5,
    height: HEADER_HEIGHT + 20 * 10,
    ...over,
  });

describe('Viewport', () => {
  it('lists only the lines on screen', () => {
    const v = make();
    expect(v.visibleColumns.map((b) => b.index)).toEqual([0, 1, 2, 3, 4]);
    expect(v.visibleRows).toHaveLength(10);
  });

  it('keeps frozen lines in place while the rest scrolls', () => {
    const v = make({ frozenRows: 2, scrollY: 100 });
    expect(v.visibleRows.slice(0, 3).map((b) => [b.index, b.frozen])).toEqual([
      [0, true],
      [1, true],
      [7, false],
    ]);
    expect(v.rowY(7)).toBe(HEADER_HEIGHT + 40);
    expect(v.rowAt(HEADER_HEIGHT + 45)).toBe(7);
    expect(v.rowAt(HEADER_HEIGHT + 5)).toBe(0);
  });

  it('hit-tests headers and cells under zoom', () => {
    const v = make({ zoom: 2 });
    expect(v.hitTest(10, 10)).toEqual({ zone: 'corner' });
    expect(v.hitTest(HEADER_WIDTH + 130, 5)).toEqual({ zone: 'columnHeader', column: 1 });
    expect(v.hitTest(HEADER_WIDTH + 1, HEADER_HEIGHT + 41)).toEqual({ zone: 'cell', row: 1, column: 0 });
  });

  it('scrolls just far enough to reveal a cell', () => {
    const v = make();
    expect(v.scrollToReveal('y', 3)).toBe(0);
    expect(v.scrollToReveal('y', 15)).toBe(16 * 20 - 200);
    expect(make({ scrollY: 400 }).scrollToReveal('y', 2)).toBe(40);
  });
});
