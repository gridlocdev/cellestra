import { describe, expect, it } from 'vitest';
import {
  CellRange,
  CellValue,
  EditCommands,
  EditRefusedError,
  OdsWriter,
  Workbook,
  type EditCommand,
} from '@cellestra/core';
import { WorkbookSession } from '../src/session/workbookSession.ts';

const refusing: EditCommand = {
  label: 'Sort',
  apply() {
    throw new EditRefusedError('Cells that are merged can’t be sorted.');
  },
  undo() {
    throw new Error('never applied');
  },
};

describe('WorkbookSession.execute', () => {
  it('reports a refused edit and leaves the history alone', () => {
    const session = new WorkbookSession();
    const told: string[] = [];
    session.onRefused = (m) => told.push(m);
    session.execute(refusing);
    expect(told).toEqual(['Cells that are merged can’t be sorted.']);
    expect(session.canUndo.value).toBe(false);
    expect(session.modified.value).toBe(false);
  });

  it('still throws other errors', () => {
    const session = new WorkbookSession();
    const broken: EditCommand = {
      ...refusing,
      apply: () => {
        throw new Error('bug');
      },
    };
    expect(() => session.execute(broken)).toThrow('bug');
  });
});

describe('WorkbookSession.cellsChangedSince', () => {
  it('lists edited and recalculated cells, and is null after an edit it cannot follow', () => {
    const session = new WorkbookSession();
    const sheet = session.activeSheet.peek();
    sheet.setFormula(0, 1, 'A1*2');
    session.recalculateAll();
    const start = session.version.peek();
    session.execute(EditCommands.setInput(sheet.name, 0, 0, '5', 'en-US'));
    expect(sheet.getValue(0, 1)).toEqual(CellValue.number(10));
    expect(session.cellsChangedSince(sheet.name, start)).toEqual([
      { row: 0, column: 0 },
      { row: 0, column: 1 },
    ]);
    expect(session.cellsChangedSince('Other', start)).toEqual([]);
    const wb = session.workbook.peek();
    const bold = (range: string) =>
      EditCommands.applyStyle(wb, sheet.name, CellRange.parse(range), (s) => ({
        ...s,
        font: { ...s.font, bold: true },
      }));
    const before = session.version.peek();
    session.execute(bold('C3'));
    expect(session.cellsChangedSince(sheet.name, before)).toEqual([{ row: 2, column: 2 }]);
    // A whole-column style reports no cells: the grid has to measure everything again.
    session.execute(bold('D1:D1048576'));
    expect(session.cellsChangedSince(sheet.name, start)).toBeNull();
  });
});

describe('WorkbookSession.deleteSheet', () => {
  it('turns references into #REF!, can be undone, and never re-links to a new sheet of that name', () => {
    const session = new WorkbookSession();
    session.addSheet(); // Sheet2
    const wb = session.workbook.value;
    wb.worksheets[1]!.setValue(0, 0, CellValue.number(42));
    session.setActiveSheet(0);
    session.execute(EditCommands.setInput('Sheet1', 0, 0, '=Sheet2.A1*2'));
    expect(wb.worksheets[0]!.getValue(0, 0)).toEqual(CellValue.number(84));

    session.deleteSheet(1);
    expect(session.sheetNames.value).toEqual(['Sheet1']);
    expect(wb.worksheets[0]!.find(0, 0)?.formula).toContain('#REF!');

    session.undo();
    expect(session.sheetNames.value).toEqual(['Sheet1', 'Sheet2']);
    expect(wb.worksheets[0]!.getValue(0, 0)).toEqual(CellValue.number(84));

    session.redo();
    session.addSheet(); // a new, empty "Sheet2"
    expect(session.sheetNames.value).toEqual(['Sheet1', 'Sheet2']);
    expect(wb.worksheets[0]!.getValue(0, 0).kind).toBe('error');
  });
});

describe('WorkbookSession.load', () => {
  function bytes(mode: 'automatic' | 'manual'): Uint8Array {
    const wb = Workbook.createNew();
    wb.settings.calculationMode = mode;
    const sheet = wb.worksheets[0]!;
    sheet.setValue(0, 0, CellValue.number(2));
    sheet.setFormula(0, 1, 'A1*2', CellValue.number(99)); // a cached value the file disagrees with
    return OdsWriter.write(wb, null);
  }

  it('recalculates on open in automatic mode', () => {
    const session = new WorkbookSession();
    session.load(bytes('automatic'), null, 'a.ods');
    expect(session.activeSheet.value.getValue(0, 1)).toEqual(CellValue.number(4));
  });

  it('keeps the cached values in manual mode until a recalculation is asked for', () => {
    const session = new WorkbookSession();
    session.load(bytes('manual'), null, 'm.ods');
    expect(session.activeSheet.value.getValue(0, 1)).toEqual(CellValue.number(99));
    session.recalculateVolatile();
    expect(session.activeSheet.value.getValue(0, 1)).toEqual(CellValue.number(99));
    session.recalculateAll();
    expect(session.activeSheet.value.getValue(0, 1)).toEqual(CellValue.number(4));
  });
});
