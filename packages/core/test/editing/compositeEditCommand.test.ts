import { describe, expect, it } from 'vitest';
import { ActionCommand, EditResult, type ChangedCell } from '../../src/editing/editCommand.ts';
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

  it('combines hundreds of thousands of changed cells', () => {
    // A sort of A1:D40000 reports about 160,000 cells per step; spreading that many arguments
    // overflows the call stack.
    const workbook = Workbook.createNew();
    const cell: ChangedCell = { sheet: 'Sheet1', row: 0, column: 0 };
    const many = new Array<ChangedCell>(300_000).fill(cell);
    const step = new ActionCommand(
      'Big',
      () => EditResult.cells(many),
      () => EditResult.cells(many),
    );
    const composite = EditCommands.composite('Two big steps', [step, step]);

    expect(composite.apply(workbook).changedCells).toHaveLength(600_000);
    expect(composite.undo(workbook).changedCells).toHaveLength(600_000);
  });

  it('takes every step back when combining their results fails', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    const broken = new ActionCommand(
      'Broken result',
      () => ({
        ...EditResult.none,
        get changedCells(): readonly ChangedCell[] {
          throw new Error('unreadable');
        },
      }),
      () => EditResult.none,
    );
    const composite = EditCommands.composite('Edit', [
      EditCommands.setValue(sheet.name, 0, 0, CellValue.number(1)),
      broken,
    ]);

    expect(() => composite.apply(workbook)).toThrow('unreadable');
    expect(sheet.getValue(0, 0).kind).toBe('empty');
  });
});
