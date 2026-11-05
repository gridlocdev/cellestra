import { describe, expect, it } from 'vitest';
import { DataRegion } from '../../src/data/dataRegion.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

function sheetWith(cells: [row: number, column: number][]) {
  const wb = new Workbook();
  const sheet = wb.addSheet('S');
  for (const [r, c] of cells) sheet.setValue(r, c, CellValue.number(1));
  return sheet;
}

describe('DataRegion.currentRegion', () => {
  it('is the single cell when the cell is empty', () => {
    expect(DataRegion.currentRegion(sheetWith([]), 3, 2).toString()).toBe('C4');
  });

  it('grows over a block and stops at blank rows and columns', () => {
    const block: [number, number][] = [];
    for (let r = 1; r <= 4; r++) for (let c = 1; c <= 3; c++) block.push([r, c]);
    block.push([6, 1], [1, 5]); // separated by a blank row / column
    expect(DataRegion.currentRegion(sheetWith(block), 2, 2).toString()).toBe('B2:D5');
  });

  it('follows an L shape that only connects after growing the other way', () => {
    // A1:A3 down, then C3 touches B3 → B3 joins via row 3, then C1 via column C.
    const sheet = sheetWith([
      [0, 0],
      [1, 0],
      [2, 0],
      [2, 1],
      [2, 2],
      [0, 2],
    ]);
    expect(DataRegion.currentRegion(sheet, 0, 0).toString()).toBe('A1:C3');
  });

  it('stays linear on a long table', () => {
    const cells: [number, number][] = [];
    for (let r = 0; r < 50_000; r++) cells.push([r, 0], [r, 1]);
    const sheet = sheetWith(cells);
    const t = performance.now();
    expect(DataRegion.currentRegion(sheet, 0, 0).toString()).toBe('A1:B50000');
    expect(performance.now() - t).toBeLessThan(1000);
  });
});
