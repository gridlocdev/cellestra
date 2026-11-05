import { describe, expect, it } from 'vitest';
import { HiddenLines } from '../../src/editing/hiddenLines.ts';
import { SheetDimensions } from '../../src/model/sheetDimensions.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';

// Which hidden lines an Unhide reaches (PRD §8.14): a hidden line can't be selected, so the selection
// the user can make sits beside the lines they want back.
const fresh = (): Worksheet => Workbook.createNew().activeSheet;

function hideRows(sheet: Worksheet, ...rows: number[]): void {
  for (const row of rows) sheet.getOrCreateRow(row).hidden = true;
}

describe('HiddenLines', () => {
  it('nothing hidden means nothing to unhide', () => {
    const sheet = fresh();
    expect(HiddenLines.rowsToUnhide(sheet, 0, 10)).toEqual([]);
    expect(HiddenLines.columnsToUnhide(sheet, 0, 10)).toEqual([]);
  });

  it('hidden rows inside the selection come back as one run', () => {
    const sheet = fresh();
    hideRows(sheet, 4, 5, 6);
    expect(HiddenLines.rowsToUnhide(sheet, 3, 7)).toEqual([[4, 6]]);
  });

  it('the row above a hidden run reaches the whole run', () => {
    const sheet = fresh();
    hideRows(sheet, 4, 5, 6);
    expect(HiddenLines.rowsToUnhide(sheet, 3, 3)).toEqual([[4, 6]]);
    expect(HiddenLines.rowsToUnhide(sheet, 7, 7)).toEqual([[4, 6]]); // …and from below
  });

  it('the reach stops at the first visible line', () => {
    const sheet = fresh();
    hideRows(sheet, 4, 5);
    hideRows(sheet, 8, 9); // a second run, with visible rows 6..7 between
    expect(HiddenLines.rowsToUnhide(sheet, 3, 3)).toEqual([[4, 5]]);
    expect(HiddenLines.rowsToUnhide(sheet, 3, 10)).toEqual([
      [4, 5],
      [8, 9],
    ]);
  });

  it('a selection nowhere near a hidden run finds nothing', () => {
    const sheet = fresh();
    hideRows(sheet, 4, 5, 6);
    expect(HiddenLines.rowsToUnhide(sheet, 20, 25)).toEqual([]);
  });

  it('a whole-sheet selection reports only the hidden runs', () => {
    const sheet = fresh();
    hideRows(sheet, 4, 5, 6);
    expect(HiddenLines.rowsToUnhide(sheet, 0, SheetDimensions.maxRows - 1)).toEqual([[4, 6]]);
  });

  it('columns behave the same as rows', () => {
    const sheet = fresh();
    sheet.getOrCreateColumn(2).hidden = true;
    sheet.getOrCreateColumn(3).hidden = true;
    expect(HiddenLines.columnsToUnhide(sheet, 1, 1)).toEqual([[2, 3]]);
    expect(HiddenLines.columnsToUnhide(sheet, 4, 4)).toEqual([[2, 3]]);
    expect(HiddenLines.columnsToUnhide(sheet, 8, 9)).toEqual([]);
  });

  it('the reach never runs off the edge of the sheet', () => {
    const sheet = fresh();
    hideRows(sheet, 0, 1);
    expect(HiddenLines.rowsToUnhide(sheet, 2, 2)).toEqual([[0, 1]]);
    expect(HiddenLines.rowsToUnhide(sheet, 0, 0)).toEqual([[0, 1]]);
  });
});
