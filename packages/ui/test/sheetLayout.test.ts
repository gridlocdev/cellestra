import { CellValue, Workbook } from '@cellestra/core';
import { describe, expect, it } from 'vitest';
import { contentEnd, sheetLayout } from '../src/grid/sheetLayout.ts';

describe('sheet layout', () => {
  it('gives hidden lines no size and grows rows to their content', () => {
    const sheet = Workbook.createNew().activeSheet;
    sheet.getOrCreateColumn(1).width = 100;
    sheet.getOrCreateColumn(2).hidden = true;
    const layout = sheetLayout(sheet, new Map([[4, 45]]));
    expect(layout.columns.size(1)).toBe(100);
    expect(layout.columns.size(2)).toBe(0);
    expect(layout.rows.size(4)).toBe(45);
    expect(layout.rows.size(5)).toBe(20);
    expect(layout.lastColumn).toBe(2);
  });

  it('reaches past cells to merges and laid-out lines', () => {
    const sheet = Workbook.createNew().activeSheet;
    sheet.setValue(3, 1, CellValue.number(1));
    sheet.getOrCreateColumn(6).width = 90;
    expect(contentEnd(sheet, sheetLayout(sheet, new Map()))).toEqual({ bottom: 3, right: 6 });
  });
});
