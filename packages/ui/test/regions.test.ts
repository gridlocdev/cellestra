import { describe, expect, it } from 'vitest';
import { nextRegion, type RegionNode } from '../src/components/regions.ts';

/** Nodes laid out in a line: a node at `at` follows every node before it; a region holds its children. */
function node(at: number, children: number[] = []): RegionNode & { at: number } {
  return {
    at,
    contains: (other) =>
      other !== null &&
      ((other as { at: number }).at === at || children.includes((other as { at: number }).at)),
    compareDocumentPosition: (other) => ((other as { at: number }).at > at ? 4 : 2),
  };
}

describe('nextRegion', () => {
  const ribbon = node(10, [11, 12]);
  const grid = node(30, [31]);
  const sheets = node(40, [41, 42]);
  const regions = [ribbon, grid, sheets];

  it('steps from the region that has focus and wraps both ways', () => {
    expect(nextRegion(regions, node(11), 1)).toBe(grid);
    expect(nextRegion(regions, node(31), 1)).toBe(sheets);
    expect(nextRegion(regions, node(42), 1)).toBe(ribbon);
    expect(nextRegion(regions, node(12), -1)).toBe(sheets);
    expect(nextRegion(regions, node(30), -1)).toBe(ribbon);
  });

  it('from outside every region, goes to the nearest one in that direction', () => {
    const formulaBar = node(20);
    expect(nextRegion(regions, formulaBar, 1)).toBe(grid);
    expect(nextRegion(regions, formulaBar, -1)).toBe(ribbon);
    expect(nextRegion(regions, node(50), 1)).toBe(ribbon);
    expect(nextRegion(regions, node(5), -1)).toBe(sheets);
  });

  it('starts at an end with nothing focused, and has nowhere to go with no regions', () => {
    expect(nextRegion(regions, null, 1)).toBe(ribbon);
    expect(nextRegion(regions, null, -1)).toBe(sheets);
    expect(nextRegion([], node(1), 1)).toBeUndefined();
  });
});
