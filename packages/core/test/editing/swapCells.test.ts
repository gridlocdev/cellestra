import { describe, expect, it } from 'vitest';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { SwapCellsCommand } from '../../src/editing/swapCellsCommand.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const R = (a1: string): CellRange => CellRange.parse(a1);

function setup() {
  const workbook = Workbook.createNew();
  return { history: new EditHistory(workbook), sheet: workbook.activeSheet };
}

const numAt = (sheet: Worksheet, row: number, column: number): number =>
  CellValue.asNumber(sheet.find(row, column)!.value);

describe('SwapCellsCommand', () => {
  it("swap exchanges two cells' contents", () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(10));
    sheet.setValue(0, 1, CellValue.number(20));

    history.execute(new SwapCellsCommand(sheet.name, R('A1'), 0, 1));

    expect(numAt(sheet, 0, 0)).toBe(20);
    expect(numAt(sheet, 0, 1)).toBe(10);

    history.undo();
    expect(numAt(sheet, 0, 0)).toBe(10);
    expect(numAt(sheet, 0, 1)).toBe(20);
  });

  it('swap with an empty destination is a plain move', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(10));

    history.execute(new SwapCellsCommand(sheet.name, R('A1'), 0, 1));

    expect(sheet.find(0, 0)).toBeFalsy();
    expect(numAt(sheet, 0, 1)).toBe(10);
  });

  it('swap makes references follow both ways without hijack', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(10));
    sheet.setValue(0, 1, CellValue.number(20));
    sheet.setFormula(0, 2, 'A1', CellValue.number(10));
    sheet.setFormula(0, 3, 'B1', CellValue.number(20));

    history.execute(new SwapCellsCommand(sheet.name, R('A1'), 0, 1));

    expect(sheet.find(0, 2)!.formula).toBe('B1');
    expect(sheet.find(0, 3)!.formula).toBe('A1');

    history.undo();
    expect(sheet.find(0, 2)!.formula).toBe('A1');
    expect(sheet.find(0, 3)!.formula).toBe('B1');
  });

  it('swap is its own inverse', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(10));
    sheet.setValue(0, 1, CellValue.number(20));
    sheet.setFormula(0, 2, 'A1+B1', CellValue.number(30));

    history.execute(new SwapCellsCommand(sheet.name, R('A1'), 0, 1));
    history.execute(new SwapCellsCommand(sheet.name, R('A1'), 0, 1));

    expect(numAt(sheet, 0, 0)).toBe(10);
    expect(numAt(sheet, 0, 1)).toBe(20);
    expect(sheet.find(0, 2)!.formula).toBe('A1+B1');
  });

  it('swapping two blocks carries intra-block references along', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.setFormula(1, 0, 'A1+1', CellValue.number(2));

    history.execute(new SwapCellsCommand(sheet.name, R('A1:A2'), 0, 2));

    expect(sheet.find(1, 2)!.formula).toBe('C1+1');
  });
});
