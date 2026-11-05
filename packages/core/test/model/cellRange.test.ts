import { describe, expect, it } from 'vitest';
import { CellRange } from '../../src/model/cellRange.ts';

describe('CellRange', () => {
  it('parse of a single cell is one by one', () => {
    const range = CellRange.parse('C5');
    expect(range.isSingleCell).toBe(true);
    expect(range.cellCount).toBe(1);
    expect(range.left).toBe(2);
    expect(range.top).toBe(4);
  });

  it('parse normalizes corner order', () => {
    const range = CellRange.parse('C5:A1');
    expect([range.left, range.top, range.right, range.bottom]).toEqual([0, 0, 2, 4]);
    expect(range.columnCount).toBe(3);
    expect(range.rowCount).toBe(5);
    expect(range.cellCount).toBe(15);
  });

  it('a sheet qualifier on the start applies to the range', () => {
    const range = CellRange.parse('Sheet1.A1:B2');
    expect(range.sheet).toBe('Sheet1');
    expect(range.toString()).toBe('Sheet1.A1:B2');
  });

  it.each([
    'Sheet1.A1:Sheet2.B2', // conflicting sheet on the end corner
    'A1:Sheet2.B2', // lone sheet qualifier on the end corner
  ])('a cross-sheet range is rejected: %s', (text) => {
    expect(CellRange.tryParse(text)).toBeNull();
  });

  it('a matching sheet on both corners parses', () => {
    expect(CellRange.tryParse('Sheet1.A1:Sheet1.B2')?.sheet).toBe('Sheet1');
  });

  it('contains and intersects', () => {
    const a = CellRange.parse('B2:D4');
    expect(a.contains(2, 2)).toBe(true);
    expect(a.contains(0, 0)).toBe(false);
    expect(a.intersects(CellRange.parse('D4:F6'))).toBe(true);
    expect(a.intersects(CellRange.parse('E5:F6'))).toBe(false);
  });

  it('positions enumerates row-major', () => {
    expect([...CellRange.parse('A1:B2').positions()]).toEqual([
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
    ]);
  });
});
