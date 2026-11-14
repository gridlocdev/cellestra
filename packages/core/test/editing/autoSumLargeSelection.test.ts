import { describe, expect, it } from 'vitest';
import { AutoSumPlanner } from '../../src/editing/autoSumPlanner.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { SheetDimensions } from '../../src/model/sheetDimensions.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const LAST_ROW = SheetDimensions.maxRows - 1;
const LAST_COLUMN = SheetDimensions.maxColumns - 1;

describe('AutoSum on huge selections', () => {
  it('only reads the used part of a whole-sheet selection', () => {
    const sheet = Workbook.createNew().activeSheet;
    sheet.setValue(0, 1, CellValue.number(1));
    sheet.setValue(1, 1, CellValue.text('x'));
    const selection = CellRange.fromBounds(0, 0, LAST_ROW - 1, LAST_COLUMN);
    expect(AutoSumPlanner.plan(sheet, selection, 0, 0)).toEqual([
      { targetRow: LAST_ROW, targetColumn: 1, formula: `=SUM(B1:B${LAST_ROW})` },
    ]);
  });

  it('falls back to the active cell when a huge selection holds no numbers', () => {
    const sheet = Workbook.createNew().activeSheet;
    sheet.setValue(0, 0, CellValue.number(4));
    sheet.setValue(5, 3, CellValue.text('label'));
    const selection = CellRange.fromBounds(0, 1, LAST_ROW, LAST_COLUMN);
    expect(AutoSumPlanner.plan(sheet, selection, 1, 0)).toEqual([
      { targetRow: 1, targetColumn: 0, formula: '=SUM(A1)' },
    ]);
  });
});
