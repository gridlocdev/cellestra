import { describe, expect, it } from 'vitest';
import {
  CellRange,
  CellValue,
  EditCommands,
  EditRefusedError,
  OdsWriter,
  StructuralEditCommand,
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
    // A whole-column style reports its column: the cells stored in it are the ones to measure again.
    sheet.setValue(5, 3, CellValue.text('x'));
    sheet.setValue(5, 4, CellValue.text('y'));
    const styled = session.version.peek();
    session.execute(bold('D1:D1048576'));
    expect(session.cellsChangedSince(sheet.name, styled)).toEqual([{ row: 5, column: 3 }]);
    // So does a resized row or column.
    const resized = session.version.peek();
    session.execute(EditCommands.setColumnWidth(sheet.name, 4, 200));
    session.execute(EditCommands.setRowHeight(sheet.name, 5, 40));
    expect(session.cellsChangedSince(sheet.name, resized)).toEqual([
      { row: 5, column: 4 },
      { row: 5, column: 3 },
      { row: 5, column: 4 },
    ]);
    // An edit that reports nothing can't be followed: the grid measures everything again.
    session.execute(EditCommands.setFrozenPanes(sheet.name, 1, 0));
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

describe('WorkbookSession status', () => {
  it('narrates edits, undo and redo, and forgets them when another document opens', () => {
    const session = new WorkbookSession();
    const sheet = session.activeSheet.peek().name;
    expect(session.lastAction.value).toBe('');
    session.execute(EditCommands.setInput(sheet, 1, 1, '5', 'en-US'));
    expect(session.lastAction.value).toBe('Edited cell B2');
    session.execute(
      EditCommands.setBlock(sheet, 0, 0, [
        [CellValue.number(1), CellValue.number(2)],
        [CellValue.number(3), CellValue.number(4)],
      ]),
    );
    expect(session.lastAction.value).toBe('Pasted A1:B2');
    session.undo();
    expect(session.lastAction.value).toBe('Undid Paste');
    session.redo();
    expect(session.lastAction.value).toBe('Redid Paste');
    session.newDocument();
    expect(session.lastAction.value).toBe('');
  });

  it('records when the document was saved', () => {
    const session = new WorkbookSession();
    const at = new Date(2026, 10, 14, 14, 30);
    session.now = () => at;
    expect(session.savedAt.value).toBeNull();
    session.execute(EditCommands.setInput(session.activeSheet.peek().name, 0, 0, '1', 'en-US'));
    session.markSaved({ store: 'browser', id: 'x', name: 'x.ods', version: null });
    expect(session.savedAt.value).toBe(at);
    expect(session.modified.value).toBe(false);
    session.newDocument();
    expect(session.savedAt.value).toBeNull();
  });
});

describe('WorkbookSession.setCalculationMode', () => {
  it('holds results in manual mode, and catches up when switched back to automatic', () => {
    const session = new WorkbookSession();
    const sheet = session.activeSheet.peek();
    session.execute(EditCommands.setInput(sheet.name, 0, 1, '=A1*2', 'en-US'));
    session.markSaved({ store: 'browser', id: 'x', name: 'x.ods', version: null });

    session.setCalculationMode('manual');
    expect(session.calculationMode.value).toBe('manual');
    expect(session.workbook.peek().settings.calculationMode).toBe('manual');
    expect(session.modified.value).toBe(true);
    expect(session.lastAction.value).toBe('Calculation: Manual');

    session.execute(EditCommands.setInput(sheet.name, 0, 0, '5', 'en-US'));
    expect(sheet.getValue(0, 1)).toEqual(CellValue.number(0));
    session.recalculateAll();
    expect(sheet.getValue(0, 1)).toEqual(CellValue.number(10));
    expect(session.lastAction.value).toBe('Recalculated');

    session.execute(EditCommands.setInput(sheet.name, 0, 0, '7', 'en-US'));
    expect(sheet.getValue(0, 1)).toEqual(CellValue.number(10));
    session.setCalculationMode('automatic');
    expect(sheet.getValue(0, 1)).toEqual(CellValue.number(14));
    expect(session.calculationMode.value).toBe('automatic');
  });
});

describe('WorkbookSession manual calculation and structural edits', () => {
  /** A1 = 1, A2 = 2, B1 = SUM(A1:A2), C1 = A2*10; then manual calculation. */
  function manualSession() {
    const session = new WorkbookSession();
    const sheet = session.activeSheet.peek();
    const set = (row: number, column: number, text: string) =>
      session.execute(EditCommands.setInput(sheet.name, row, column, text, 'en-US'));
    set(0, 0, '1');
    set(1, 0, '2');
    set(0, 1, '=SUM(A1:A2)');
    set(0, 2, '=A2*10');
    session.setCalculationMode('manual');
    return { session, sheet, set };
  }

  it('inserting and deleting rows adjusts references but waits for F9 to recalculate', () => {
    const { session, sheet, set } = manualSession();
    set(1, 0, '5');
    expect(sheet.getValue(0, 1)).toEqual(CellValue.number(3)); // stale until F9

    session.execute(new StructuralEditCommand(sheet.name, 'row', 'insert', 1, 1));
    expect(sheet.find(0, 1)?.formula).toBe('SUM(A1:A3)');
    expect(sheet.find(0, 2)?.formula).toBe('A3*10');
    expect(sheet.getValue(0, 1)).toEqual(CellValue.number(3));
    expect(sheet.getValue(0, 2)).toEqual(CellValue.number(20));

    set(1, 0, '100');
    expect(sheet.getValue(0, 1)).toEqual(CellValue.number(3));
    session.recalculateAll();
    expect(sheet.getValue(0, 1)).toEqual(CellValue.number(106));
    expect(sheet.getValue(0, 2)).toEqual(CellValue.number(50));

    session.execute(new StructuralEditCommand(sheet.name, 'row', 'delete', 2, 1));
    expect(sheet.find(0, 2)?.formula).toContain('#REF!');
    expect(sheet.getValue(0, 1)).toEqual(CellValue.number(106));
    expect(sheet.getValue(0, 2)).toEqual(CellValue.number(50));
    session.undo();
    expect(sheet.find(0, 2)?.formula).toBe('A3*10');
    session.recalculateAll();
    expect(sheet.getValue(0, 1)).toEqual(CellValue.number(106));
    expect(sheet.getValue(0, 2)).toEqual(CellValue.number(50));
  });

  it('inserting columns and sheets, renaming, and named ranges wait too', () => {
    const { session, sheet, set } = manualSession();
    set(1, 0, '7');
    session.execute(new StructuralEditCommand(sheet.name, 'column', 'insert', 0, 1));
    expect(sheet.find(0, 2)?.formula).toBe('SUM(B1:B2)');
    expect(sheet.getValue(0, 2)).toEqual(CellValue.number(3));

    session.addSheet();
    session.setActiveSheet(0);
    expect(session.renameSheet(0, 'Data')).toBeNull();
    session.namesChanged();
    expect(sheet.getValue(0, 2)).toEqual(CellValue.number(3));
    expect(sheet.getValue(0, 3)).toEqual(CellValue.number(20));

    session.recalculateAll();
    expect(sheet.getValue(0, 2)).toEqual(CellValue.number(8));
    expect(sheet.getValue(0, 3)).toEqual(CellValue.number(70));
  });

  it('recalculates structural edits as before in automatic mode', () => {
    const { session, sheet, set } = manualSession();
    session.setCalculationMode('automatic');
    session.execute(new StructuralEditCommand(sheet.name, 'row', 'insert', 1, 1));
    set(1, 0, '4');
    expect(sheet.getValue(0, 1)).toEqual(CellValue.number(7));
  });
});

describe('WorkbookSession and hidden rows', () => {
  it('recalculates SUBTOTAL when rows are hidden or shown, and logs the cell it changed', () => {
    const session = new WorkbookSession();
    const sheet = session.activeSheet.peek();
    [10, 20, 30].forEach((n, row) => sheet.setValue(row, 0, CellValue.number(n)));
    sheet.setFormula(4, 0, 'SUBTOTAL(109, A1:A3)');
    session.recalculateAll();
    expect(sheet.getValue(4, 0)).toEqual(CellValue.number(60));
    const before = session.version.peek();
    session.execute(EditCommands.setRowsHidden(sheet.name, 1, 1, true));
    expect(sheet.getValue(4, 0)).toEqual(CellValue.number(40));
    expect(session.cellsChangedSince(sheet.name, before)).toContainEqual({ row: 4, column: 0 });
    session.undo();
    expect(sheet.getValue(4, 0)).toEqual(CellValue.number(60));
  });
});
