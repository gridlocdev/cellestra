import { describe, expect, it } from 'vitest';
import { EditCommands } from '../../src/editing/editCommands.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

describe('CompositeEditCommand', () => {
  it('applies all parts and undoes them as one step', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    const composite = EditCommands.composite('Two edits', [
      EditCommands.setValue(sheet.name, 0, 0, CellValue.number(1)),
      EditCommands.setValue(sheet.name, 0, 1, CellValue.number(2)),
    ]);

    composite.apply(workbook);
    expect(CellValue.asNumber(sheet.getValue(0, 0))).toBe(1);
    expect(CellValue.asNumber(sheet.getValue(0, 1))).toBe(2);

    composite.undo(workbook);
    expect(sheet.getValue(0, 0).kind).toBe('empty');
    expect(sheet.getValue(0, 1).kind).toBe('empty');
  });
});
