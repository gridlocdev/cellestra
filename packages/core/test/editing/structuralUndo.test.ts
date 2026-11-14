import { describe, expect, it } from 'vitest';
import { StructuralEditCommand } from '../../src/editing/structuralEditCommand.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

function book() {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  const other = workbook.addSheet('Other');
  for (let r = 0; r < 5; r++) sheet.setValue(r, 0, CellValue.number(r + 1));
  sheet.setFormula(4, 1, 'SUM(A1:A5)', CellValue.number(15));
  sheet.setFormula(0, 2, 'A2*2', CellValue.number(4));
  other.setFormula(0, 0, 'Sheet1.A2+Sheet1.A5', CellValue.number(7));
  return { workbook, sheet, other };
}

describe('Structural undo', () => {
  it('moves the existing cells instead of copying the sheet, and puts them back on undo', () => {
    const { workbook, sheet } = book();
    const a5 = sheet.find(4, 0)!;
    const command = new StructuralEditCommand(sheet.name, 'row', 'delete', 1, 1);
    command.apply(workbook);
    expect(sheet.find(3, 0)).toBe(a5);
    command.undo(workbook);
    expect(sheet.find(4, 0)).toBe(a5);
  });

  it('restores deleted cells and every rewritten formula', () => {
    const { workbook, sheet, other } = book();
    const command = new StructuralEditCommand(sheet.name, 'row', 'delete', 1, 1);
    command.apply(workbook);
    expect(sheet.find(3, 1)?.formula).toBe('SUM(A1:A4)');
    expect(sheet.find(0, 2)?.formula).toBe('#REF!*2');
    expect(other.find(0, 0)?.formula).toBe('#REF!+Sheet1.A4');
    command.undo(workbook);
    expect(CellValue.asNumber(sheet.getValue(1, 0))).toBe(2);
    expect(sheet.find(4, 1)?.formula).toBe('SUM(A1:A5)');
    expect(sheet.find(0, 2)?.formula).toBe('A2*2');
    expect(other.find(0, 0)?.formula).toBe('Sheet1.A2+Sheet1.A5');
    expect(sheet.cellCount).toBe(7);
  });

  it('undoes an insert, dropping the formatting the new lines inherited', () => {
    const { workbook, sheet } = book();
    sheet.getOrCreate(0, 3).styleId = 1;
    const command = new StructuralEditCommand(sheet.name, 'row', 'insert', 1, 2);
    command.apply(workbook);
    expect(sheet.find(1, 3)?.styleId).toBe(1);
    expect(sheet.find(6, 1)?.formula).toBe('SUM(A1:A7)');
    command.undo(workbook);
    expect(sheet.find(1, 3)).toBeUndefined();
    expect(sheet.find(4, 1)?.formula).toBe('SUM(A1:A5)');
    expect(CellValue.asNumber(sheet.getValue(1, 0))).toBe(2);
  });

  it('undoes a column delete', () => {
    const { workbook, sheet } = book();
    const command = new StructuralEditCommand(sheet.name, 'column', 'delete', 0, 1);
    command.apply(workbook);
    expect(sheet.find(4, 0)?.formula).toBe('SUM(#REF!)');
    command.undo(workbook);
    expect(sheet.find(4, 1)?.formula).toBe('SUM(A1:A5)');
    expect(CellValue.asNumber(sheet.getValue(2, 0))).toBe(3);
  });
});
