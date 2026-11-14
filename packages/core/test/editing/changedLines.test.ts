import { describe, expect, it } from 'vitest';
import { DataEditCommands } from '../../src/editing/dataEditCommands.ts';
import { combineResults, EditResult, type EditCommand } from '../../src/editing/editCommand.ts';
import { EditCommands } from '../../src/editing/editCommands.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { FormatBlock } from '../../src/editing/formatBlock.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { SheetDimensions } from '../../src/model/sheetDimensions.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { CellStyle } from '../../src/styling/cellStyle.ts';
import { CellValue } from '../../src/values/cellValue.ts';

// Edits report the whole rows and columns whose cells may now look or measure differently, so a
// view re-measures only those instead of the whole sheet.

const R = (a1: string): CellRange => CellRange.parse(a1);
const big = (s: CellStyle): CellStyle => ({ ...s, font: { ...s.font, sizePoints: 24 } });
const lines = (sheet: string, first: number, last = first) => [{ sheet, first, last }];
const LAST_ROW = SheetDimensions.maxRows - 1;
const LAST_COLUMN = SheetDimensions.maxColumns - 1;

function setup() {
  const workbook = Workbook.createNew();
  const history = new EditHistory(workbook);
  const name = workbook.activeSheet.name;
  /** apply, undo and redo all report the same lines. */
  const each = (command: EditCommand): EditResult[] => [
    history.execute(command),
    history.undo(),
    history.redo(),
  ];
  return { workbook, history, name, each };
}

describe('changed rows and columns', () => {
  it('a value edit reports cells only', () => {
    const { name, history } = setup();
    const result = history.execute(EditCommands.setValue(name, 0, 0, CellValue.number(1)));
    expect(result.changedCells.length).toBe(1);
    expect(result.changedRows).toEqual([]);
    expect(result.changedColumns).toEqual([]);
  });

  it('resizing a column reports it', () => {
    const { name, each } = setup();
    for (const r of each(EditCommands.setColumnWidth(name, 3, 200))) {
      expect(r.changedColumns).toEqual(lines(name, 3));
      expect(r.changedRows).toEqual([]);
    }
  });

  it('resizing a row reports it', () => {
    const { name, each } = setup();
    for (const r of each(EditCommands.setRowHeight(name, 4, 40)))
      expect(r.changedRows).toEqual(lines(name, 4));
  });

  it('hiding lines reports them', () => {
    const { name, each } = setup();
    for (const r of each(EditCommands.setColumnsHidden(name, 1, 3, true)))
      expect(r.changedColumns).toEqual(lines(name, 1, 3));
    for (const r of each(EditCommands.setRowsHidden(name, 5, 9, true)))
      expect(r.changedRows).toEqual(lines(name, 5, 9));
  });

  it('styling whole columns reports the columns', () => {
    const { workbook, name, each } = setup();
    for (const r of each(
      EditCommands.applyStyle(workbook, name, CellRange.fromBounds(0, 1, LAST_ROW, 2), big),
    )) {
      expect(r.changedColumns).toEqual(lines(name, 1, 2));
      expect(r.changedRows).toEqual([]);
    }
  });

  it('styling whole rows reports the rows', () => {
    const { workbook, name, each } = setup();
    for (const r of each(
      EditCommands.applyStyle(workbook, name, CellRange.fromBounds(1, 0, 3, LAST_COLUMN), big),
    )) {
      expect(r.changedRows).toEqual(lines(name, 1, 3));
      expect(r.changedColumns).toEqual([]);
    }
  });

  it('styling the whole sheet reports every column', () => {
    const { workbook, name, each } = setup();
    const all = CellRange.fromBounds(0, 0, LAST_ROW, LAST_COLUMN);
    for (const r of each(EditCommands.applyStyle(workbook, name, all, big)))
      expect(r.changedColumns).toEqual(lines(name, 0, LAST_COLUMN));
  });

  it('merging and unmerging report the merged rows', () => {
    const { workbook, name, each } = setup();
    for (const r of each(EditCommands.mergeCells(name, R('B2:C5'))))
      expect(r.changedRows).toEqual(lines(name, 1, 4));
    workbook.activeSheet.setValue(1, 1, CellValue.text('x'));
    for (const r of each(EditCommands.unmergeCells(name, 2, 2)))
      expect(r.changedRows).toEqual(lines(name, 1, 4));
  });

  it('clearing structure reports the rows it covers and the merges it removed', () => {
    const { workbook, name, each } = setup();
    workbook.activeSheet.setValue(0, 0, CellValue.text('x'));
    new EditHistory(workbook).execute(EditCommands.mergeCells(name, R('A1:A10')));
    for (const r of each(DataEditCommands.clearRangeStructure(name, R('A3:B4'))))
      expect(r.changedRows).toEqual([...lines(name, 2, 3), ...lines(name, 0, 9)]);
    for (const r of each(DataEditCommands.clearRangeStructure(name, CellRange.fromBounds(0, 2, LAST_ROW, 3))))
      expect(r.changedColumns).toEqual(lines(name, 2, 3));
  });

  it('painting formats over part of a merge reports the merge rows', () => {
    const { workbook, name, each } = setup();
    const sheet = workbook.activeSheet;
    new EditHistory(workbook).execute(EditCommands.mergeCells(name, R('C1:C8')));
    const stamp = FormatBlock.capture(workbook, sheet, R('A1:A2'));
    new EditHistory(workbook).execute(EditCommands.mergeCells(name, R('A1:A2')));
    const withMerge = FormatBlock.capture(workbook, sheet, R('A1:A2'));
    for (const r of each(withMerge.applyTo(name, R('C3:C4'))))
      expect(r.changedRows).toContainEqual(lines(name, 0, 7)[0]);
    expect(stamp.applyTo(name, R('E1')).apply(workbook).changedRows).toEqual([]);
  });

  it('resetting the sheet layout reports every column', () => {
    const { name, each } = setup();
    for (const r of each(DataEditCommands.resetSheetLayout(name)))
      expect(r.changedColumns).toEqual(lines(name, 0, LAST_COLUMN));
  });

  it('a composite reports the lines of every step', () => {
    const { name, history } = setup();
    const result = history.execute(
      EditCommands.composite('Resize', [
        EditCommands.setColumnWidth(name, 1, 90),
        EditCommands.setColumnWidth(name, 2, 90),
        EditCommands.setRowHeight(name, 0, 30),
      ]),
    );
    expect(result.changedColumns).toEqual([...lines(name, 1), ...lines(name, 2)]);
    expect(result.changedRows).toEqual(lines(name, 0));
  });

  it('combineResults keeps the lines', () => {
    const combined = combineResults([
      EditResult.columns('S', 1, 2),
      EditResult.rows('S', 3, 3),
      EditResult.none,
    ]);
    expect(combined.changedColumns).toEqual(lines('S', 1, 2));
    expect(combined.changedRows).toEqual(lines('S', 3));
    expect(combined.structural).toBe(false);
  });
});
