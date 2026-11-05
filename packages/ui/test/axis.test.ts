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

describe('Axis hidden runs', () => {
  const count = 1_048_576;
  const hiddenTail = Array.from({ length: count - 10 }, (_, k) => [k + 10, 0] as [number, number]);

  it('jumps a hidden tail in one step', () => {
    const axis = new Axis(count, 20, hiddenTail);
    const t = performance.now();
    for (let n = 0; n < 1000; n++) {
      expect(axis.step(9, 1)).toBe(9);
      expect(axis.firstVisibleFrom(500_000)).toBe(9);
      expect(axis.nextVisible(10)).toBe(count);
    }
    expect(performance.now() - t).toBeLessThan(200);
  });

  it('steps over a hidden run in the middle and stays put with nothing visible', () => {
    const axis = new Axis(20, 20, [
      [0, 0],
      [3, 0],
      [4, 0],
      [5, 30],
      [6, 0],
    ]);
    expect(axis.step(2, 1)).toBe(5);
    expect(axis.step(5, 1)).toBe(7);
    expect(axis.step(7, -1)).toBe(5);
    expect(axis.step(5, -1)).toBe(2);
    expect(axis.step(1, -1)).toBe(1); // row 0 is hidden: nowhere to go
    expect(axis.prevVisible(4)).toBe(2);
    expect(axis.lastVisibleBefore(0)).toBe(1);
  });
});
