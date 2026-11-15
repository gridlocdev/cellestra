import { describe, expect, it } from 'vitest';
import { filterSpec } from '../../src/data/filterEngine.ts';
import { dataValidation, table as makeTable } from '../../src/data/types.ts';
import { DataEditCommands } from '../../src/editing/dataEditCommands.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellStyles, Fills } from '../../src/styling/cellStyle.ts';
import { Argb } from '../../src/styling/color.ts';
import { StylePool } from '../../src/styling/stylePool.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const R = (a1: string): CellRange => CellRange.parse(a1);

function setup() {
  const workbook = Workbook.createNew();
  return { workbook, history: new EditHistory(workbook), sheet: workbook.activeSheet };
}

const text = (v: CellValue): string | null => (v.kind === 'text' ? v.value : null);

describe('DataEditCommands', () => {
  it('sort is undoable', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('H'));
    sheet.setValue(1, 0, CellValue.number(3));
    sheet.setValue(2, 0, CellValue.number(1));
    sheet.setValue(3, 0, CellValue.number(2));

    history.execute(DataEditCommands.sortByColumn(sheet.name, R('A1:A4'), 0, true));
    expect(CellValue.asNumber(sheet.getValue(1, 0))).toBe(1);
    expect(CellValue.asNumber(sheet.getValue(3, 0))).toBe(3);

    history.undo();
    expect(CellValue.asNumber(sheet.getValue(1, 0))).toBe(3);
  });

  it('header column styles the first column like the header and undoes', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('Team'));
    sheet.setValue(0, 1, CellValue.text('Pts'));
    sheet.setValue(1, 0, CellValue.text('A'));
    sheet.setValue(1, 1, CellValue.number(3));
    sheet.setValue(2, 0, CellValue.text('B'));
    sheet.setValue(2, 1, CellValue.number(1));
    const table = makeTable('T', R('A1:B3'), { hasHeader: true });
    sheet.addTable(table);

    history.execute(DataEditCommands.setTableHeaderColumn(workbook, sheet.name, table, true));

    expect(sheet.tables[0]!.hasHeaderColumn).toBe(true);
    const a2 = workbook.styles.get(sheet.find(1, 0)!.styleId);
    expect(a2.font.bold).toBe(true);
    expect(Fills.hasBackground(a2.fill)).toBe(false);
    expect(workbook.styles.get(sheet.find(1, 1)!.styleId).font.bold).toBe(false);

    history.undo();
    expect(sheet.tables[0]!.hasHeaderColumn ?? false).toBe(false);
    expect(workbook.styles.get(sheet.find(1, 0)!.styleId).font.bold).toBe(false);
  });

  it('table header commands bake no colours into cells', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('Site'));
    sheet.setValue(1, 0, CellValue.text('All Sites'));
    const table = makeTable('T', R('A1:B2'), { hasHeader: true });
    history.execute(DataEditCommands.createStyledTable(workbook, sheet.name, table));
    history.execute(DataEditCommands.setTableHeaderColumn(workbook, sheet.name, table, true));

    for (const [row, column] of [
      [0, 0],
      [0, 1],
      [1, 0],
    ] as const) {
      const style = workbook.styles.get(sheet.find(row, column)!.styleId);
      expect(style.font.bold).toBe(true);
      expect(Fills.hasBackground(style.fill)).toBe(false);
    }
  });

  it('banded rows toggle flips the flag and undoes', () => {
    const { history, sheet } = setup();
    const table = makeTable('T', R('A1:B4'), { hasHeader: true });
    sheet.addTable(table);
    expect(sheet.tables[0]!.hasBandedRows ?? false).toBe(false);

    history.execute(DataEditCommands.setTableBandedRows(sheet.name, table, true));
    expect(sheet.tables[0]!.hasBandedRows).toBe(true);
    expect(sheet.find(1, 0)).toBeFalsy();

    history.undo();
    expect(sheet.tables[0]!.hasBandedRows ?? false).toBe(false);
  });

  it('header column can be cleared back to default', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('Team'));
    sheet.setValue(1, 0, CellValue.text('A'));
    const table = makeTable('T', R('A1:A2'), { hasHeader: true });
    sheet.addTable(table);

    history.execute(DataEditCommands.setTableHeaderColumn(workbook, sheet.name, table, true));
    const on = sheet.tables[0]!;
    history.execute(DataEditCommands.setTableHeaderColumn(workbook, sheet.name, on, false));

    expect(sheet.tables[0]!.hasHeaderColumn ?? false).toBe(false);
    const a2 = workbook.styles.get(sheet.find(1, 0)!.styleId);
    expect(a2.font.bold).toBe(false);
    expect(Fills.hasBackground(a2.fill)).toBe(false);
  });

  it('filter hides rows and undo restores', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('Score'));
    sheet.setValue(1, 0, CellValue.number(10));
    sheet.setValue(2, 0, CellValue.number(90));

    const spec = filterSpec(new Map([[0, { condition: '>40' }]]));
    history.execute(DataEditCommands.applyFilter(sheet.name, R('A1:A3'), spec, 'en-US'));
    expect(sheet.isRowHidden(1)).toBe(true);
    expect(sheet.isRowHidden(2)).toBe(false);

    history.undo();
    expect(sheet.isRowHidden(1)).toBe(false);
  });

  it('filter reports the rows it hid or showed, on apply and on undo', () => {
    const { workbook, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('Score'));
    for (const [row, n] of [
      [1, 10],
      [2, 90],
      [3, 20],
      [4, 95],
    ] as const)
      sheet.setValue(row, 0, CellValue.number(n));
    const spec = filterSpec(new Map([[0, { condition: '>40' }]]));
    const filter = DataEditCommands.applyFilter(sheet.name, R('A1:A5'), spec, 'en-US');
    const rows = [{ sheet: sheet.name, first: 1, last: 3 }];
    expect(filter.apply(workbook).changedRows).toEqual(rows);
    expect(filter.undo(workbook).changedRows).toEqual(rows);
    const nothing = DataEditCommands.applyFilter(sheet.name, R('A1:A5'), filterSpec(new Map()), 'en-US');
    expect(nothing.apply(workbook).changedRows).toEqual([]);
  });

  it('resetSheetLayout clears row/column fills, hidden and filter, and undo restores', () => {
    const { workbook, history, sheet } = setup();
    const fillId = workbook.styles.intern({
      ...CellStyles.default,
      fill: Fills.solid(Argb.fromRgb(255, 0, 0)),
    });
    sheet.getOrCreateColumn(1).styleId = fillId;
    sheet.getOrCreateRow(2).manuallyHidden = true;
    sheet.activeFilter = R('A1:B3');
    sheet.activeFilterSpec = filterSpec(new Map([[0, { condition: '>0' }]]));

    history.execute(DataEditCommands.resetSheetLayout(sheet.name));

    expect(sheet.getOrCreateColumn(1).styleId).toBe(0);
    expect(sheet.isRowHidden(2)).toBe(false);
    expect(sheet.activeFilter).toBeNull();
    expect(sheet.activeFilterSpec).toBeNull();

    history.undo();

    expect(sheet.getOrCreateColumn(1).styleId).toBe(fillId);
    expect(sheet.isRowHidden(2)).toBe(true);
    expect(sheet.activeFilter!.equals(R('A1:B3'))).toBe(true);
    expect(sheet.activeFilterSpec).not.toBeNull();
  });

  it('remove duplicates is undoable and reports count', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('H'));
    sheet.setValue(1, 0, CellValue.text('a'));
    sheet.setValue(2, 0, CellValue.text('a'));
    sheet.setValue(3, 0, CellValue.text('b'));

    const { command, removedCount } = DataEditCommands.removeDuplicates(sheet.name, R('A1:A4'), [0], true);
    history.execute(command);

    expect(removedCount()).toBe(1);
    expect(text(sheet.getValue(2, 0))).toBe('b');

    history.undo();
    expect(text(sheet.getValue(2, 0))).toBe('a');
  });

  it('remove duplicates moves formulas with their rows', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('H'));
    ['a', 'a', 'b'].forEach((v, i) => {
      sheet.setValue(i + 1, 0, CellValue.text(v));
      sheet.setFormula(i + 1, 1, `A${i + 2}`);
    });
    history.execute(DataEditCommands.removeDuplicates(sheet.name, R('A1:B4'), [0], true).command);
    expect(sheet.find(2, 1)!.formula).toBe('A3');
    expect(sheet.find(3, 1)).toBeUndefined();
    history.undo();
    expect(sheet.find(3, 1)!.formula).toBe('A4');
  });

  it('remove duplicates over whole columns only touches the used rows', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('H'));
    sheet.setValue(1, 0, CellValue.text('a'));
    sheet.setValue(2, 0, CellValue.text('a'));
    const whole = CellRange.fromBounds(0, 0, 1_048_575, 0);
    const { command, removedCount } = DataEditCommands.removeDuplicates(sheet.name, whole, [0], true);
    history.execute(command);
    expect(removedCount()).toBe(1);
    expect(sheet.find(2, 0)).toBeUndefined();
    history.undo();
    expect(text(sheet.getValue(2, 0))).toBe('a');
  });

  it('undoing a sort over a filter shows the rows it showed before', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('H'));
    [5, 50, 30].forEach((n, i) => sheet.setValue(i + 1, 0, CellValue.number(n)));
    const spec = filterSpec(new Map([[0, { condition: '>40' }]]));
    history.execute(DataEditCommands.applySortAndFilter(sheet.name, R('A1:A4'), null, spec, 'en-US'));
    const hidden = () => [1, 2, 3].map((r) => sheet.isRowHidden(r));
    expect(hidden()).toEqual([true, false, true]);

    history.execute(DataEditCommands.sortByColumn(sheet.name, R('A1:A4'), 0, false, true, 'en-US'));
    expect(hidden()).toEqual([false, true, true]);

    history.undo();
    expect(sheet.getValue(1, 0)).toEqual(CellValue.number(5));
    expect(hidden()).toEqual([true, false, true]);
  });

  it('text to columns is undoable', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('a,b,c'));

    history.execute(DataEditCommands.textToColumns(sheet, R('A1:A1'), ','));
    expect(text(sheet.getValue(0, 1))).toBe('b');

    history.undo();
    expect(sheet.getValue(0, 1).kind).toBe('empty');
    expect(text(sheet.getValue(0, 0))).toBe('a,b,c');
  });

  it('create table is undoable', () => {
    const { history, sheet } = setup();
    const table = makeTable('Data', R('A1:B3'));

    history.execute(DataEditCommands.createTable(sheet.name, table));
    expect(sheet.findTable(0, 0)).toBeTruthy();

    history.undo();
    expect(sheet.findTable(0, 0)).toBeFalsy();
  });

  it('create table styles header row with contrasting default', () => {
    const { workbook, history, sheet } = setup();
    const table = makeTable('Data', R('A1:B3'), { hasHeader: true });

    history.execute(DataEditCommands.createStyledTable(workbook, sheet.name, table));

    for (const column of [0, 1]) {
      const header = workbook.styles.get(sheet.getOrCreate(0, column).styleId);
      expect(header.font.bold).toBe(true);
      expect(Fills.hasBackground(header.fill)).toBe(false);
    }
    expect(workbook.styles.get(sheet.getOrCreate(1, 0).styleId).font.bold).toBe(false);

    history.undo();
    expect(sheet.findTable(0, 0)).toBeFalsy();
    expect(sheet.getOrCreate(0, 0).styleId).toBe(StylePool.defaultId);
  });

  it('create table prefills empty header cells left to right', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(0, 1, CellValue.text('Name'));
    const table = makeTable('Data', R('A1:C3'), { hasHeader: true });

    history.execute(DataEditCommands.createStyledTable(workbook, sheet.name, table));

    expect(text(sheet.getValue(0, 0))).toBe('Column 1');
    expect(text(sheet.getValue(0, 1))).toBe('Name');
    expect(text(sheet.getValue(0, 2))).toBe('Column 3');
    expect(workbook.styles.get(sheet.getOrCreate(0, 0).styleId).font.bold).toBe(true);

    history.undo();
    expect(sheet.getValue(0, 0).kind).toBe('empty');
    expect(text(sheet.getValue(0, 1))).toBe('Name');
  });

  it('create table without header skips styling', () => {
    const { workbook, history, sheet } = setup();
    const table = makeTable('Data', R('A1:B3'), { hasHeader: false });

    history.execute(DataEditCommands.createStyledTable(workbook, sheet.name, table));

    expect(sheet.getOrCreate(0, 0).styleId).toBe(StylePool.defaultId);
  });

  it('expand table for edit below grows the range and is undoable', () => {
    const { workbook, history, sheet } = setup();
    const table = makeTable('Data', R('A1:B3'), { hasHeader: true });
    history.execute(DataEditCommands.createStyledTable(workbook, sheet.name, table));

    const command = DataEditCommands.expandTableForEdit(workbook, sheet.name, table, 3, 0);
    expect(command).not.toBeNull();
    history.execute(command!);

    expect(sheet.tables).toHaveLength(1);
    expect(sheet.tables[0]!.range.equals(R('A1:B4'))).toBe(true);

    history.undo();
    expect(sheet.tables).toHaveLength(1);
    expect(sheet.tables[0]!.range.equals(R('A1:B3'))).toBe(true);
  });

  it('expand table for edit right grows range and styles new header', () => {
    const { workbook, history, sheet } = setup();
    const table = makeTable('Data', R('A1:B3'), { hasHeader: true });
    history.execute(DataEditCommands.createStyledTable(workbook, sheet.name, table));

    const command = DataEditCommands.expandTableForEdit(workbook, sheet.name, table, 0, 2);
    expect(command).not.toBeNull();
    history.execute(command!);

    expect(sheet.tables).toHaveLength(1);
    expect(sheet.tables[0]!.range.equals(R('A1:C3'))).toBe(true);
    const header = workbook.styles.get(sheet.getOrCreate(0, 2).styleId);
    expect(header.font.bold).toBe(true);
    expect(Fills.hasBackground(header.fill)).toBe(false);

    history.undo();
    expect(sheet.tables[0]!.range.equals(R('A1:B3'))).toBe(true);
    expect(sheet.getOrCreate(0, 2).styleId).toBe(StylePool.defaultId);
  });

  it('expand table for edit below carries the header column onto the new row', () => {
    const { workbook, history, sheet } = setup();
    const table = makeTable('Data', R('A1:B3'), { hasHeader: true, hasHeaderColumn: true });
    history.execute(DataEditCommands.createStyledTable(workbook, sheet.name, table));

    const command = DataEditCommands.expandTableForEdit(workbook, sheet.name, table, 3, 1);
    expect(command).not.toBeNull();
    history.execute(command!);

    const header = workbook.styles.get(sheet.getOrCreate(3, 0).styleId);
    expect(header.font.bold).toBe(true);
    expect(Fills.hasBackground(header.fill)).toBe(false);

    history.undo();
    expect(sheet.getOrCreate(3, 0).styleId).toBe(StylePool.defaultId);
  });

  it('expand table for edit below leaves the first cell plain without a header column', () => {
    const { workbook, history, sheet } = setup();
    const table = makeTable('Data', R('A1:B3'), { hasHeader: true });
    history.execute(DataEditCommands.createStyledTable(workbook, sheet.name, table));

    history.execute(DataEditCommands.expandTableForEdit(workbook, sheet.name, table, 3, 1)!);

    expect(sheet.getOrCreate(3, 0).styleId).toBe(StylePool.defaultId);
  });

  it('expand table for edit returns null when not adjacent', () => {
    const { workbook, sheet } = setup();
    const table = makeTable('Data', R('A1:B3'), { hasHeader: true });

    expect(DataEditCommands.expandTableForEdit(workbook, sheet.name, table, 10, 10)).toBeNull();
  });

  it('set validation is undoable', () => {
    const { history, sheet } = setup();
    const rule = dataValidation('list', R('A1'), { listValues: ['Yes', 'No'] });

    history.execute(DataEditCommands.setValidation(sheet.name, rule));
    expect(sheet.findValidation(0, 0)).toBeTruthy();

    history.undo();
    expect(sheet.findValidation(0, 0)).toBeFalsy();
  });

  it('replace validation keeps the parts of old rules outside the range and undoes', () => {
    const { history, sheet } = setup();
    const old = dataValidation('list', R('A1:A10'), { listValues: ['a'] });
    history.execute(DataEditCommands.setValidation(sheet.name, old));

    const rule = dataValidation('list', R('A4:A5'), { listValues: ['b'] });
    history.execute(DataEditCommands.replaceValidation(sheet.name, R('A4:A5'), rule));
    const ranges = sheet.validations.map((v) => `${v.target.toString()}=${v.listValues!.join()}`);
    expect(ranges.sort()).toEqual(['A1:A3=a', 'A4:A5=b', 'A6:A10=a']);

    history.execute(DataEditCommands.replaceValidation(sheet.name, R('A1:A10'), null));
    expect(sheet.validations).toEqual([]);

    history.undo();
    history.undo();
    expect(sheet.validations).toEqual([old]);
  });
});
