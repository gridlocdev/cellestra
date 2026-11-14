import { describe, expect, it } from 'vitest';
import { MAX_COLUMN, MAX_ROW, Selection } from '../src/session/selection.ts';

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

describe('Selection.cycle skip', () => {
  it('skips cells the predicate rejects', () => {
    const s = new Selection();
    s.selectRange({ top: 0, left: 0, bottom: 3, right: 0 });
    const hidden = new Set([1, 2]);
    s.cycle(1, 0, (row) => hidden.has(row));
    expect(s.active.value).toEqual({ row: 3, column: 0 });
    s.cycle(1, 0, (row) => hidden.has(row));
    expect(s.active.value).toEqual({ row: 0, column: 0 });
  });
});

describe('Selection.selectLines', () => {
  it('widens every range to whole rows or columns and keeps the active cell', () => {
    const s = new Selection();
    s.select(2, 3);
    s.select(4, 5, { extend: true });
    s.selectLines('rows');
    expect(s.state.value.ranges).toEqual([{ top: 2, left: 0, bottom: 4, right: MAX_COLUMN }]);
    expect(s.active.value).toEqual({ row: 2, column: 3 });
    s.selectLines('columns');
    expect(s.state.value.ranges).toEqual([{ top: 0, left: 0, bottom: MAX_ROW, right: MAX_COLUMN }]);
  });

  it('widens each range of a discontiguous selection', () => {
    const s = new Selection();
    s.select(1, 1);
    s.select(6, 2, { add: true });
    s.selectLines('columns');
    expect(s.state.value.ranges).toEqual([
      { top: 0, left: 1, bottom: MAX_ROW, right: 1 },
      { top: 0, left: 2, bottom: MAX_ROW, right: 2 },
    ]);
  });
});

describe('Selection.collapse', () => {
  it('shrinks a range or a discontiguous selection to the active cell', () => {
    const s = new Selection();
    s.select(2, 2);
    s.select(5, 4, { extend: true });
    expect(s.collapse()).toBe(true);
    expect(s.state.value.ranges).toEqual([{ top: 2, left: 2, bottom: 2, right: 2 }]);
    s.select(7, 7, { add: true });
    expect(s.collapse()).toBe(true);
    expect(s.state.value.ranges).toEqual([{ top: 7, left: 7, bottom: 7, right: 7 }]);
  });

  it('does nothing for a single cell, and keeps a merged active cell whole', () => {
    const s = new Selection();
    s.select(3, 3);
    expect(s.collapse()).toBe(false);
    const merge = { top: 3, left: 3, bottom: 4, right: 5 };
    s.selectRange({ top: 0, left: 0, bottom: 9, right: 9 }, { row: 3, column: 3 });
    expect(s.collapse(merge)).toBe(true);
    expect(s.range.value).toEqual(merge);
    expect(s.collapse(merge)).toBe(false);
  });
});
