import { describe, expect, it } from 'vitest';
import { table } from '../../src/data/types.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { StructuralEditCommand, type StructuralAxis } from '../../src/editing/structuralEditCommand.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';

// Inserting/deleting rows or columns that intersect a table resizes the table along the edited
// axis (PRD §8.10); undo restores the original range.

const R = (a1: string): CellRange => CellRange.parse(a1);

function setup() {
  const workbook = Workbook.createNew();
  return { history: new EditHistory(workbook), sheet: workbook.activeSheet };
}

const del = (history: EditHistory, sheet: Worksheet, axis: StructuralAxis, at: number, count: number) =>
  history.execute(new StructuralEditCommand(sheet.name, axis, 'delete', at, count));

const insert = (history: EditHistory, sheet: Worksheet, axis: StructuralAxis, at: number, count: number) =>
  history.execute(new StructuralEditCommand(sheet.name, axis, 'insert', at, count));

function expectSingleTable(sheet: Worksheet, a1: string): void {
  expect(sheet.tables).toHaveLength(1);
  expect(sheet.tables[0]!.range.equals(R(a1))).toBe(true);
}

describe('Structural edits and tables', () => {
  it('deleting rows inside a table shrinks it', () => {
    const { history, sheet } = setup();
    sheet.addTable(table('T', R('A2:B10')));

    del(history, sheet, 'row', 4, 3);

    expectSingleTable(sheet, 'A2:B7');

    history.undo();
    expectSingleTable(sheet, 'A2:B10');
  });

  it('deleting columns inside a table shrinks it', () => {
    const { history, sheet } = setup();
    sheet.addTable(table('T', R('B1:F5')));

    del(history, sheet, 'column', 2, 2);

    expectSingleTable(sheet, 'B1:D5');
  });

  it('deleting rows above a table shifts it up unchanged in size', () => {
    const { history, sheet } = setup();
    sheet.addTable(table('T', R('A5:B8')));

    del(history, sheet, 'row', 0, 2);

    expectSingleTable(sheet, 'A3:B6');
  });

  it('deleting rows below a table leaves it untouched', () => {
    const { history, sheet } = setup();
    sheet.addTable(table('T', R('A1:B4')));

    del(history, sheet, 'row', 6, 2);

    expectSingleTable(sheet, 'A1:B4');
  });

  it('deleting every row of a table removes it', () => {
    const { history, sheet } = setup();
    sheet.addTable(table('T', R('A3:B5')));

    del(history, sheet, 'row', 2, 4);

    expect(sheet.tables).toHaveLength(0);

    history.undo();
    expectSingleTable(sheet, 'A3:B5');
  });

  it('inserting rows inside a table grows it', () => {
    const { history, sheet } = setup();
    sheet.addTable(table('T', R('A2:B6')));

    insert(history, sheet, 'row', 4, 2);

    expectSingleTable(sheet, 'A2:B8');
  });
});
