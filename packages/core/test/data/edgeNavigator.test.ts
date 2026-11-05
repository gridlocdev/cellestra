import { describe, expect, it } from 'vitest';
import { EdgeNavigator } from '../../src/data/edgeNavigator.ts';
import { SheetDimensions } from '../../src/model/sheetDimensions.ts';
import { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';

/** Excel-style ⌘+Arrow "jump to edge" navigation (PRD parity). */
function sheetWith(...cells: [number, number, number][]): Worksheet {
  const sheet = new Worksheet('Sheet1');
  for (const [row, column, value] of cells) sheet.setValue(row, column, CellValue.number(value));
  return sheet;
}

describe('EdgeNavigator', () => {
  it('within a contiguous run jumps to the last filled cell', () => {
    const sheet = sheetWith([0, 0, 1], [1, 0, 2], [2, 0, 3], [3, 0, 4]);
    expect(EdgeNavigator.findEdge(sheet, 0, 0, 1, 0)).toEqual([3, 0]);
  });

  it('facing a gap jumps across to the next filled cell', () => {
    const sheet = sheetWith([0, 0, 1], [4, 0, 5]);
    expect(EdgeNavigator.findEdge(sheet, 0, 0, 1, 0)).toEqual([4, 0]);
  });

  it('from an empty cell jumps to the next filled cell', () => {
    expect(EdgeNavigator.findEdge(sheetWith([3, 0, 1]), 0, 0, 1, 0)).toEqual([3, 0]);
  });

  it('from the last filled cell jumps to the sheet edge', () => {
    const sheet = sheetWith([0, 0, 1], [1, 0, 2]);
    expect(EdgeNavigator.findEdge(sheet, 1, 0, 1, 0)).toEqual([SheetDimensions.maxRows - 1, 0]);
  });

  it('empty sheet jumps to the boundary', () => {
    const sheet = new Worksheet('Sheet1');
    expect(EdgeNavigator.findEdge(sheet, 0, 0, 0, 1)).toEqual([0, SheetDimensions.maxColumns - 1]);
    expect(EdgeNavigator.findEdge(sheet, 0, 0, -1, 0)).toEqual([0, 0]);
  });

  it('works horizontally to the right', () => {
    const sheet = sheetWith([0, 0, 1], [0, 1, 2], [0, 2, 3]);
    expect(EdgeNavigator.findEdge(sheet, 0, 0, 0, 1)).toEqual([0, 2]);
  });
});
