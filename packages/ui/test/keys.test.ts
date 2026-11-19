import { describe, expect, it } from 'vitest';
import { afterTab, enterOrigin, startsComposition } from '../src/grid/keys.ts';

const key = (key: string, keyCode = 0, isComposing = false) => ({ key, keyCode, isComposing });

describe('startsComposition', () => {
  it('is true for dead keys and IME keys', () => {
    expect(startsComposition(key('Dead', 222))).toBe(true);
    expect(startsComposition(key('Process', 229))).toBe(true);
    expect(startsComposition(key('a', 229))).toBe(true);
    expect(startsComposition(key('n', 78, true))).toBe(true);
  });

  it('is false for plain characters and commands', () => {
    expect(startsComposition(key('a', 65))).toBe(false);
    expect(startsComposition(key('Enter', 13))).toBe(false);
    expect(startsComposition(key('ArrowDown', 40))).toBe(false);
  });
});

describe('Tab then Enter', () => {
  it('returns to the column where the Tab run started, one row down', () => {
    let run = afterTab(null, { row: 1, column: 2 }, { row: 1, column: 3 });
    run = afterTab(run, { row: 1, column: 3 }, { row: 1, column: 4 });
    expect(enterOrigin(run, { row: 1, column: 4 })).toEqual({ row: 1, column: 2 });
  });

  it('keeps the start through Shift+Tab', () => {
    let run = afterTab(null, { row: 0, column: 0 }, { row: 0, column: 1 });
    run = afterTab(run, { row: 0, column: 1 }, { row: 0, column: 2 });
    run = afterTab(run, { row: 0, column: 2 }, { row: 0, column: 1 });
    expect(enterOrigin(run, { row: 0, column: 1 })).toEqual({ row: 0, column: 0 });
  });

  it('goes straight down after Shift+Tab left of where the run started', () => {
    let run = afterTab(null, { row: 1, column: 1 }, { row: 1, column: 0 });
    expect(enterOrigin(run, { row: 1, column: 0 })).toEqual({ row: 1, column: 0 });
    run = afterTab(null, { row: 1, column: 2 }, { row: 1, column: 3 });
    run = afterTab(run, { row: 1, column: 3 }, { row: 1, column: 2 });
    run = afterTab(run, { row: 1, column: 2 }, { row: 1, column: 1 });
    expect(enterOrigin(run, { row: 1, column: 1 })).toEqual({ row: 1, column: 1 });
  });

  it('starts a new run once the active cell moved away from the last Tab', () => {
    const run = afterTab(null, { row: 1, column: 2 }, { row: 1, column: 3 });
    expect(enterOrigin(run, { row: 5, column: 3 })).toEqual({ row: 5, column: 3 });
    expect(afterTab(run, { row: 1, column: 6 }, { row: 1, column: 7 })).toEqual({
      row: 1,
      startColumn: 6,
      column: 7,
    });
  });

  it('ends at a Tab that changes rows', () => {
    expect(afterTab(null, { row: 1, column: 9 }, { row: 2, column: 0 })).toBeNull();
    expect(enterOrigin(null, { row: 2, column: 0 })).toEqual({ row: 2, column: 0 });
  });
});
