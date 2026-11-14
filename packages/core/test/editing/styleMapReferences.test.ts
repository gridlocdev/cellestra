import { describe, expect, it } from 'vitest';
import { RenameSheetCommand } from '../../src/editing/renameSheetCommand.ts';
import { StructuralEditCommand } from '../../src/editing/structuralEditCommand.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellStyles, type CellStyleMap } from '../../src/styling/cellStyle.ts';
import { CellValue } from '../../src/values/cellValue.ts';

/** Sheet1 with B3 styled by a conditional map based on B3 that compares it with D1 on Other. */
function book() {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  workbook.addSheet('Other');
  const map: CellStyleMap = {
    condition: 'is-true-formula([.B3]>[Other.$D$1])',
    applyStyleName: 'Good',
    applyStyle: null,
    baseCellAddress: '$Sheet1.$B$3',
  };
  const id = workbook.styles.intern({ ...CellStyles.default, conditionalMaps: [map] });
  sheet.setValue(2, 1, CellValue.number(5));
  sheet.getOrCreate(2, 1).styleId = id;
  const mapOf = () => workbook.styles.get(sheet.find(...cellOf())!.styleId).conditionalMaps![0]!;
  let at: [number, number] = [2, 1];
  const cellOf = () => at;
  const moveTo = (row: number, column: number) => (at = [row, column]);
  return { workbook, mapOf, moveTo };
}

describe('conditional maps follow structural edits', () => {
  it('moves the base cell and the condition when rows go in above, and undo puts them back', () => {
    const { workbook, mapOf, moveTo } = book();
    const insert = new StructuralEditCommand('Sheet1', 'row', 'insert', 0, 2);
    insert.apply(workbook);
    moveTo(4, 1);
    expect(mapOf().baseCellAddress).toBe('$Sheet1.$B$5');
    expect(mapOf().condition).toBe('is-true-formula([.B5]>[Other.$D$1])');
    insert.undo(workbook);
    moveTo(2, 1);
    expect(mapOf().baseCellAddress).toBe('$Sheet1.$B$3');
    expect(mapOf().condition).toBe('is-true-formula([.B3]>[Other.$D$1])');
  });

  it('follows columns deleted to the left, and leaves edits on another sheet alone', () => {
    const { workbook, mapOf, moveTo } = book();
    new StructuralEditCommand('Other', 'column', 'insert', 0, 1).apply(workbook);
    expect(mapOf().baseCellAddress).toBe('$Sheet1.$B$3');
    expect(mapOf().condition).toBe('is-true-formula([.B3]>[$Other.$E$1])');
    new StructuralEditCommand('Sheet1', 'column', 'delete', 0, 1).apply(workbook);
    moveTo(2, 0);
    expect(mapOf().baseCellAddress).toBe('$Sheet1.$A$3');
    expect(mapOf().condition).toBe('is-true-formula([.A3]>[$Other.$E$1])');
  });

  it('follows a sheet rename', () => {
    const { workbook, mapOf } = book();
    new RenameSheetCommand('Sheet1', 'Data').apply(workbook);
    expect(mapOf().baseCellAddress).toBe('$Data.$B$3');
  });
});
