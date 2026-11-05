import { describe, expect, it } from 'vitest';
import { Axis } from '../src/grid/axis.ts';

describe('Axis', () => {
  const axis = new Axis(100, 20, [
    [2, 40],
    [5, 0],
    [7, 10],
  ]);

  it('offsets sum the exceptions before a line', () => {
    expect(axis.offset(0)).toBe(0);
    expect(axis.offset(2)).toBe(40);
    expect(axis.offset(3)).toBe(80);
    expect(axis.offset(6)).toBe(120); // 0..5 = 20+20+40+20+20+0
    expect(axis.offset(8)).toBe(150);
    expect(axis.total).toBe(100 * 20 + 20 - 20 - 10);
  });

  it('hit-tests to the containing line and skips hidden lines', () => {
    expect(axis.indexAt(0)).toBe(0);
    expect(axis.indexAt(39.9)).toBe(1);
    expect(axis.indexAt(40)).toBe(2);
    expect(axis.indexAt(79)).toBe(2);
    expect(axis.indexAt(120)).toBe(6); // 5 is hidden at the same offset
    expect(axis.indexAt(1e9)).toBe(99);
  });

  it('steps over hidden lines', () => {
    expect(axis.step(4, 1)).toBe(6);
    expect(axis.step(6, -1)).toBe(4);
    expect(axis.step(0, -1)).toBe(0);
    expect(axis.step(98, 5)).toBe(99);
  });

  it('handles a million rows without per-line storage', () => {
    const rows = new Axis(1_048_576, 20, [[1_000_000, 100]]);
    expect(rows.offset(1_048_576)).toBe(1_048_576 * 20 + 80);
    expect(rows.indexAt(1_000_000 * 20 + 99)).toBe(1_000_000);
    expect(rows.indexAt(1_000_000 * 20 + 100)).toBe(1_000_001);
  });
});
