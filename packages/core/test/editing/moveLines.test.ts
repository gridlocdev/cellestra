import { describe, expect, it } from 'vitest';
import { EditRefusedError } from '../../src/editing/editCommand.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { MoveLinesCommand } from '../../src/editing/moveLinesCommand.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { MergedRegion } from '../../src/model/mergedRegion.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';

/** Sheet1 A1:A5 = 1…5, B1 = SUM(A1:A5), C1 = A2*2, D1 = $A$4; Other.A1 = Sheet1.A2. */
function book() {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  const other = workbook.addSheet('Other');
  for (let r = 0; r < 5; r++) sheet.setValue(r, 0, CellValue.number(r + 1));
  sheet.setFormula(0, 1, 'SUM(A1:A5)', CellValue.number(15));
  sheet.setFormula(0, 2, 'A2*2', CellValue.number(4));
  sheet.setFormula(0, 3, '$A$4', CellValue.number(4));
  other.setFormula(0, 0, 'Sheet1.A2', CellValue.number(2));
  return { workbook, sheet, other, history: new EditHistory(workbook) };
}

const columnA = (sheet: Worksheet): number[] =>
  [0, 1, 2, 3, 4].map((r) => CellValue.asNumber(sheet.getValue(r, 0)));

describe('MoveLinesCommand', () => {
  it('moves a row up one, swapping it with the row above, and undo puts it back', () => {
    const { sheet, history } = book();
    history.execute(new MoveLinesCommand(sheet.name, 'row', 2, 2, -1));
    expect(columnA(sheet)).toEqual([1, 3, 2, 4, 5]);
    history.undo();
    expect(columnA(sheet)).toEqual([1, 2, 3, 4, 5]);
  });

  it('moves several rows down one', () => {
    const { sheet, history } = book();
    history.execute(new MoveLinesCommand(sheet.name, 'row', 1, 2, 1));
    expect(columnA(sheet)).toEqual([1, 4, 2, 3, 5]);
    history.undo();
    expect(columnA(sheet)).toEqual([1, 2, 3, 4, 5]);
  });

  it('points references at the cells where they went, absolute ones too, on every sheet', () => {
    const { sheet, other, history } = book();
    history.execute(new MoveLinesCommand(sheet.name, 'row', 1, 1, 1)); // row 2 down: rows 2 and 3 swap
    expect(sheet.find(0, 2)?.formula).toBe('A3*2');
    expect(other.find(0, 0)?.formula).toBe('Sheet1.A3');
    expect(sheet.find(0, 1)?.formula).toBe('SUM(A1:A5)'); // covers both rows: unchanged
    history.execute(new MoveLinesCommand(sheet.name, 'row', 3, 3, -1)); // row 4 up: $A$4 follows
    expect(sheet.find(0, 3)?.formula).toBe('$A$3');
    expect(sheet.find(0, 2)?.formula).toBe('A4*2'); // A3 was the swapped row
    history.undo();
    history.undo();
    expect(sheet.find(0, 2)?.formula).toBe('A2*2');
    expect(sheet.find(0, 3)?.formula).toBe('$A$4');
    expect(other.find(0, 0)?.formula).toBe('Sheet1.A2');
  });

  it('moves row heights and hidden state with the rows', () => {
    const { sheet, history } = book();
    sheet.getOrCreateRow(3).height = 40;
    history.execute(new MoveLinesCommand(sheet.name, 'row', 3, 3, -1));
    expect(sheet.getRowHeight(2)).toBe(40);
    expect(sheet.findRow(3)).toBeUndefined();
    history.undo();
    expect(sheet.getRowHeight(3)).toBe(40);
  });

  it('moves columns left and right', () => {
    const { sheet, history } = book();
    history.execute(new MoveLinesCommand(sheet.name, 'column', 0, 0, 1));
    expect(CellValue.asNumber(sheet.getValue(0, 1))).toBe(1);
    expect(sheet.find(0, 0)?.formula).toBe('SUM(B1:B5)');
    expect(sheet.find(0, 2)?.formula).toBe('B2*2');
    history.undo();
    expect(CellValue.asNumber(sheet.getValue(0, 0))).toBe(1);
    expect(sheet.find(0, 1)?.formula).toBe('SUM(A1:A5)');
  });

  it('moves a merge that sits inside the moved rows, and refuses one cut across the edge', () => {
    const { workbook, sheet, history } = book();
    sheet.addMerge(new MergedRegion(CellRange.parse('E2:F3')));
    history.execute(new MoveLinesCommand(sheet.name, 'row', 1, 2, 1));
    expect(sheet.merges[0]?.range.toString()).toBe('E3:F4');
    const split = new MoveLinesCommand(sheet.name, 'row', 2, 2, -1);
    expect(split.refusal(workbook)).toMatch(/merged/);
    expect(() => history.execute(split)).toThrow(EditRefusedError);
  });

  it('refuses to move past the edge of the sheet', () => {
    const { workbook, sheet } = book();
    expect(new MoveLinesCommand(sheet.name, 'row', 0, 1, -1).refusal(workbook)).toMatch(/edge/);
    expect(new MoveLinesCommand(sheet.name, 'column', 0, 0, -1).refusal(workbook)).toMatch(/edge/);
  });
});
