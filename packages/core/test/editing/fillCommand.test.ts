import { describe, expect, it } from 'vitest';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { FillCommand } from '../../src/editing/fillCommand.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const R = (a1: string): CellRange => CellRange.parse(a1);

function setup() {
  const workbook = Workbook.createNew();
  return { history: new EditHistory(workbook), sheet: workbook.activeSheet };
}

describe('FillCommand', () => {
  it('fill down copies a single value', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(5));
    history.execute(new FillCommand(sheet.name, R('A1'), R('A1:A4')));
    for (const row of [1, 2, 3]) expect(CellValue.asNumber(sheet.getValue(row, 0))).toBe(5);
  });

  it('fill down over whole columns stops at the end of the used cells', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(5));
    sheet.setValue(3, 2, CellValue.text('last used row'));
    history.execute(new FillCommand(sheet.name, R('A1:B1'), R('A1:B1048576')));
    for (const row of [1, 2, 3]) expect(CellValue.asNumber(sheet.getValue(row, 0))).toBe(5);
    expect(sheet.find(4, 0)).toBeUndefined();
    expect(sheet.getUsedRange()!.bottom).toBe(3);
  });

  it('fill down shifts relative references', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(10));
    sheet.setFormula(0, 1, 'A1*2', CellValue.number(20));
    history.execute(new FillCommand(sheet.name, R('B1'), R('B1:B3')));
    expect(sheet.find(1, 1)!.formula).toBe('A2*2');
    expect(sheet.find(2, 1)!.formula).toBe('A3*2');
  });

  it('fill tiles a multi-cell source pattern', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.setValue(1, 0, CellValue.number(2));
    history.execute(new FillCommand(sheet.name, R('A1:A2'), R('A1:A5')));
    expect(CellValue.asNumber(sheet.getValue(2, 0))).toBe(1); // pattern repeats (no series)
    expect(CellValue.asNumber(sheet.getValue(3, 0))).toBe(2);
    expect(CellValue.asNumber(sheet.getValue(4, 0))).toBe(1);
  });

  it('fill is undoable and leaves the source untouched', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(5));
    history.execute(new FillCommand(sheet.name, R('A1'), R('A1:A3')));
    history.undo();
    expect(CellValue.asNumber(sheet.getValue(0, 0))).toBe(5); // source kept
    expect(sheet.getValue(1, 0).kind).toBe('empty'); // fill reverted
    expect(sheet.getValue(2, 0).kind).toBe('empty');
  });
});
