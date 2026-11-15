import { describe, expect, it } from 'vitest';
import {
  CellRange,
  CellValue,
  dataValidation,
  EditHistory,
  EditRefusedError,
  namedRangeError,
  planCellEntry,
  SortFilterCommands,
  table as makeTable,
  TableCommands,
  tableNameError,
  Workbook,
} from '../../src/index.ts';

const R = (a1: string): CellRange => CellRange.parse(a1);

function setup() {
  const workbook = Workbook.createNew();
  return { workbook, history: new EditHistory(workbook), sheet: workbook.activeSheet };
}

describe('TableCommands.create', () => {
  it('uses a first row of labels as the header', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('Name'));
    sheet.setValue(1, 0, CellValue.text('Ada'));
    const plan = TableCommands.create(workbook, sheet.name, R('A1:A2'), { header: 'labels' });
    if (!plan.ok) throw new Error('refused');
    history.execute(plan.command);
    expect(sheet.tables.map((t) => [t.name, t.range.toString(), t.hasHeader])).toEqual([
      ['Table1', 'A1:A2', true],
    ]);
  });

  it('puts labels in the blank row above, or in a row inserted there', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(1, 0, CellValue.number(1));
    sheet.setValue(2, 0, CellValue.number(2));
    const above = TableCommands.create(workbook, sheet.name, R('A2:A3'), { header: 'labels' });
    if (!above.ok) throw new Error('refused');
    expect(above.table.range.toString()).toBe('A1:A3');

    const top = Workbook.createNew();
    top.activeSheet.setValue(0, 0, CellValue.number(1));
    const inserted = TableCommands.create(top, top.activeSheet.name, R('A1'), { header: 'labels' });
    if (!inserted.ok) throw new Error('refused');
    new EditHistory(top).execute(inserted.command);
    expect(top.activeSheet.getValue(1, 0)).toEqual(CellValue.number(1));
    expect(top.activeSheet.getValue(0, 0)).toEqual(CellValue.text('Column 1'));

    history.execute(above.command);
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('Column 1'));
  });

  it('refuses an overlap unless asked to replace, and names tables across the workbook', () => {
    const { workbook, history, sheet } = setup();
    workbook.addSheet('Other').addTable(makeTable('Table1', R('A1:B2')));
    sheet.addTable(makeTable('Mine', R('A1:B3')));
    const refused = TableCommands.create(workbook, sheet.name, R('B2:C4'));
    expect(refused.ok ? [] : refused.overlapping.map((t) => t.name)).toEqual(['Mine']);

    const plan = TableCommands.create(workbook, sheet.name, R('B2:C4'), { replace: true });
    if (!plan.ok) throw new Error('refused');
    history.execute(plan.command);
    expect(sheet.tables.map((t) => t.name)).toEqual(['Table2']);
    history.undo();
    expect(sheet.tables.map((t) => t.name)).toEqual(['Mine']);
  });
});

describe('SortFilterCommands', () => {
  it('sets and removes the AutoFilter in undoable steps', () => {
    const { history, sheet } = setup();
    history.execute(SortFilterCommands.setAutoFilter(sheet.name, R('A1:B5')));
    expect(sheet.activeFilter?.toString()).toBe('A1:B5');
    history.execute(SortFilterCommands.removeAutoFilter(sheet)!);
    expect(sheet.activeFilter).toBeNull();
    expect(SortFilterCommands.removeAutoFilter(sheet)).toBeNull();
    history.undo();
    expect(sheet.activeFilter?.toString()).toBe('A1:B5');
  });

  it('keeps a plain AutoFilter when its last filter and sort are cleared', () => {
    const { history, sheet } = setup();
    history.execute(SortFilterCommands.setAutoFilter(sheet.name, R('A1:B5')));
    history.execute(SortFilterCommands.sortAndFilter(sheet, { range: R('A1:B5'), table: null }, null, null));
    expect(sheet.activeFilter?.toString()).toBe('A1:B5');
  });
});

describe('planCellEntry', () => {
  it('grows a table for an entry just past it, in one step', () => {
    const { workbook, history, sheet } = setup();
    sheet.addTable(makeTable('T', R('A1:B3')));
    history.execute(planCellEntry(workbook, sheet.name, 3, 0, 'x', 'en-US').command);
    expect(sheet.tables[0]!.range.toString()).toBe('A1:B4');
    history.undo();
    expect(sheet.tables[0]!.range.toString()).toBe('A1:B3');
    expect(sheet.getValue(3, 0)).toEqual(CellValue.empty);
  });

  it('names the rule a value breaks, and never checks a formula', () => {
    const { workbook, sheet } = setup();
    const rule = dataValidation('list', R('A1'), { listValues: ['Yes', 'No'] });
    sheet.addValidation(rule);
    expect(planCellEntry(workbook, sheet.name, 0, 0, 'Maybe', 'en-US').brokenRule).toBe(rule);
    expect(planCellEntry(workbook, sheet.name, 0, 0, 'Yes', 'en-US').brokenRule).toBeNull();
    expect(planCellEntry(workbook, sheet.name, 0, 0, '=1', 'en-US').brokenRule).toBeNull();
  });
});

describe('namedRangeError', () => {
  it('refuses names that look like cells, bad characters, and duplicates', () => {
    const workbook = Workbook.createNew();
    workbook.addNamedRange({ name: 'Rate', target: R('Sheet1.A1') });
    expect(namedRangeError(workbook, 'A1')).toMatch(/cell reference/);
    expect(namedRangeError(workbook, '1st')).toMatch(/letter or underscore/);
    expect(namedRangeError(workbook, 'rate')).toMatch(/already exists/);
    expect(namedRangeError(workbook, 'Tax_Rate')).toBeNull();
  });
});

describe('Table names', () => {
  it('suggests the text above the range, then the first header cell, then TableN', () => {
    const { workbook, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('Monthly sales'));
    sheet.setValue(1, 0, CellValue.text('Month'));
    expect(TableCommands.suggestName(workbook, sheet.name, R('A2:B5'))).toBe('Monthly_sales');
    expect(TableCommands.suggestName(workbook, sheet.name, R('A1:B5'))).toBe('Monthly_sales');
    sheet.setValue(0, 0, CellValue.text('Q1')); // looks like a cell: the header's first cell instead
    expect(TableCommands.suggestName(workbook, sheet.name, R('A2:B5'))).toBe('Month');
    sheet.addTable(makeTable('Month', R('D1:D2')));
    expect(TableCommands.suggestName(workbook, sheet.name, R('A2:B5'))).toBe('Table1');
  });

  it('rejects invalid and taken names, ignoring case, but lets a table change its own case', () => {
    const { workbook, sheet } = setup();
    const sales = makeTable('Sales', R('A1:B3'));
    sheet.addTable(sales);
    workbook.addSheet('Other').addTable(makeTable('Costs', R('A1:B3')));
    workbook.addNamedRange({ name: 'Rate', target: R('Sheet1.D1') });
    expect(tableNameError(workbook, '1st')).toMatch(/start with a letter/);
    expect(tableNameError(workbook, 'My table')).toMatch(/letters, digits/);
    expect(tableNameError(workbook, 'AB12')).toMatch(/cell reference/);
    expect(tableNameError(workbook, 'costs')).toMatch(/already exists/);
    expect(tableNameError(workbook, 'RATE')).toMatch(/named range/);
    expect(tableNameError(workbook, 'SALES', sales)).toBeNull();
    expect(tableNameError(workbook, 'Revenue_2026')).toBeNull();
  });

  it('renames a table in one undoable step and refuses a bad name', () => {
    const { workbook, history, sheet } = setup();
    sheet.addTable(makeTable('Sales', R('A1:B3')));
    history.execute(TableCommands.renameTable(workbook, sheet.name, sheet.tables[0]!, ' Revenue '));
    expect(sheet.tables.map((t) => t.name)).toEqual(['Revenue']);
    history.undo();
    expect(sheet.tables.map((t) => t.name)).toEqual(['Sales']);
    history.redo();
    expect(sheet.tables.map((t) => t.name)).toEqual(['Revenue']);
    expect(() => TableCommands.renameTable(workbook, sheet.name, sheet.tables[0]!, 'A1')).toThrow(
      EditRefusedError,
    );
  });
});
