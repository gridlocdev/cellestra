import { describe, expect, it } from 'vitest';
import { Selection } from '../src/session/selection.ts';

describe('Selection', () => {
  it('extends from the anchor without moving the active cell', () => {
    const s = new Selection();
    s.select(2, 2);
    s.select(5, 4, { extend: true });
    expect(s.active.value).toEqual({ row: 2, column: 2 });
    expect(s.range.value).toEqual({ top: 2, left: 2, bottom: 5, right: 4 });
    expect(s.extentCorner).toEqual({ row: 5, column: 4 });
    s.select(0, 0, { extend: true });
    expect(s.range.value).toEqual({ top: 0, left: 0, bottom: 2, right: 2 });
  });

  it('cycles inside a range on Enter / Tab', () => {
    const s = new Selection();
    s.selectRange({ top: 0, left: 0, bottom: 1, right: 1 });
    expect(s.cycle(0, 1)).toBe(true);
    expect(s.active.value).toEqual({ row: 0, column: 1 });
    s.cycle(0, 1);
    expect(s.active.value).toEqual({ row: 1, column: 0 });
    s.cycle(1, 0);
    expect(s.active.value).toEqual({ row: 0, column: 1 });
  });

  it('adds ranges with ⌘-click', () => {
    const s = new Selection();
    s.select(0, 0);
    s.select(3, 3, { add: true });
    expect(s.state.value.ranges).toHaveLength(2);
  });
});
