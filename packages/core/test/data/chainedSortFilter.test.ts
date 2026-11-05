import { describe, expect, it } from 'vitest';
import { filterSpec, FilterEngine } from '../../src/data/filterEngine.ts';
import { sortSpec } from '../../src/data/sortEngine.ts';
import type { FilterSpec } from '../../src/data/types.ts';
import { table } from '../../src/data/types.ts';
import { DataEditCommands } from '../../src/editing/dataEditCommands.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { asText, numberAt } from './helpers.ts';

/**
 * Chaining sort and filter (the Sort & Filter dialog): context-aware value lists, a non-mutating
 * row count for the preview, and a single-undo composite restoring order and visibility together.
 */
const INV = 'en-US';

function taskSheet(): Worksheet {
  const sheet = new Worksheet('Sheet1');
  sheet.setValue(0, 0, CellValue.text('Name'));
  sheet.setValue(0, 1, CellValue.text('Status'));
  const rows: [string, string][] = [
    ['a', 'Open'],
    ['b', 'Done'],
    ['c', 'Open'],
    ['d', 'Done'],
  ];
  rows.forEach(([name, status], i) => {
    sheet.setValue(i + 1, 0, CellValue.text(name));
    sheet.setValue(i + 1, 1, CellValue.text(status));
  });
  return sheet;
}

const statusFilter = (...allowed: string[]): FilterSpec =>
  filterSpec(new Map([[1, { allowedValues: new Set(allowed) }]]));

function numberSheet(workbook: Workbook): Worksheet {
  const sheet = workbook.activeSheet;
  sheet.setValue(0, 0, CellValue.text('N'));
  sheet.setValue(0, 1, CellValue.text('S'));
  sheet.setValue(1, 0, CellValue.number(3));
  sheet.setValue(1, 1, CellValue.text('Open'));
  sheet.setValue(2, 0, CellValue.number(1));
  sheet.setValue(2, 1, CellValue.text('Done'));
  sheet.setValue(3, 0, CellValue.number(2));
  sheet.setValue(3, 1, CellValue.text('Open'));
  return sheet;
}

describe('ChainedSortFilter', () => {
  it('distinct values narrow to rows surviving the other filter', () => {
    const names = FilterEngine.distinctValues(
      taskSheet(),
      CellRange.parse('A1:B5'),
      0,
      true,
      statusFilter('Open'),
      INV,
    ).map(asText);
    expect(names).toEqual(['a', 'c']);
  });

  it("a column's own filter does not narrow its own value list", () => {
    const statuses = FilterEngine.distinctValues(
      taskSheet(),
      CellRange.parse('A1:B5'),
      1,
      true,
      statusFilter('Open'),
      INV,
    ).map(asText);
    expect(statuses).toEqual(['Open', 'Done']);
  });

  it('plain distinct overload is unfiltered', () => {
    const names = FilterEngine.distinctValues(taskSheet(), CellRange.parse('A1:B5'), 0, true).map(asText);
    expect(names).toEqual(['a', 'b', 'c', 'd']);
  });

  it('countVisible counts without hiding any row', () => {
    const sheet = taskSheet();
    const visible = FilterEngine.countVisible(sheet, CellRange.parse('A1:B5'), statusFilter('Open'), INV);
    expect(visible).toBe(2);
    expect(sheet.isRowHidden(1)).toBe(false);
    expect(sheet.isRowHidden(2)).toBe(false);
  });

  it('applySortAndFilter sorts and filters then undoes as one step', () => {
    const workbook = Workbook.createNew();
    const history = new EditHistory(workbook);
    const sheet = numberSheet(workbook);
    const range = CellRange.parse('A1:B4');
    const sort = sortSpec([{ column: 0, ascending: true }]);
    const filter = statusFilter('Open');
    sheet.activeFilter = range;
    sheet.activeFilterSpec = filter;
    sheet.activeSort = sort;

    history.execute(DataEditCommands.applySortAndFilter(sheet.name, range, sort, filter, INV));

    expect(numberAt(sheet, 1, 0)).toBe(1);
    expect(numberAt(sheet, 2, 0)).toBe(2);
    expect(numberAt(sheet, 3, 0)).toBe(3);
    expect(sheet.isRowHidden(1)).toBe(true);
    expect(sheet.isRowHidden(2)).toBe(false);

    history.undo();

    expect(numberAt(sheet, 1, 0)).toBe(3);
    expect(numberAt(sheet, 2, 0)).toBe(1);
    expect(sheet.isRowHidden(1)).toBe(false);
    expect(sheet.isRowHidden(2)).toBe(false);
  });

  it('applyTableSortAndFilter stores state on the table and undoes as one step', () => {
    const workbook = Workbook.createNew();
    const history = new EditHistory(workbook);
    const sheet = numberSheet(workbook);
    const t = table('T', CellRange.parse('A1:B4'), { hasHeader: true });
    sheet.addTable(t);
    const sort = sortSpec([{ column: 0, ascending: true }]);
    history.execute(DataEditCommands.applyTableSortAndFilter(sheet.name, t, sort, statusFilter('Open'), INV));

    expect(sheet.tables).toHaveLength(1);
    expect(sheet.tables[0]?.filter).toBeTruthy();
    expect(sheet.tables[0]?.sort).toBeTruthy();
    expect(sheet.activeFilterSpec).toBeNull();
    expect(numberAt(sheet, 1, 0)).toBe(1);
    expect(sheet.isRowHidden(1)).toBe(true);
    expect(sheet.isRowHidden(2)).toBe(false);

    history.undo();

    expect(sheet.tables).toHaveLength(1);
    expect(sheet.tables[0]?.filter ?? null).toBeNull();
    expect(sheet.tables[0]?.sort ?? null).toBeNull();
    expect(numberAt(sheet, 1, 0)).toBe(3);
    expect(sheet.isRowHidden(1)).toBe(false);
  });

  it('applySortAndFilter undo restores the sheet autofilter state', () => {
    const workbook = Workbook.createNew();
    const history = new EditHistory(workbook);
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.text('S'));
    sheet.setValue(1, 0, CellValue.text('Open'));
    sheet.setValue(2, 0, CellValue.text('Done'));
    const range = CellRange.parse('A1:A3');
    expect(sheet.activeFilter).toBeNull();

    history.execute(DataEditCommands.applySortAndFilter(sheet.name, range, null, statusFilter('Open'), INV));

    expect(sheet.activeFilter?.equals(range)).toBe(true);
    expect(sheet.activeFilterSpec).not.toBeNull();
    expect(sheet.isRowHidden(2)).toBe(true);

    history.undo();

    expect(sheet.activeFilter).toBeNull();
    expect(sheet.activeFilterSpec).toBeNull();
    expect(sheet.isRowHidden(2)).toBe(false);
  });

  it('applySortAndFilter with no filter unhides previously hidden rows', () => {
    const workbook = Workbook.createNew();
    const history = new EditHistory(workbook);
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.text('S'));
    sheet.setValue(1, 0, CellValue.text('Open'));
    sheet.setValue(2, 0, CellValue.text('Done'));
    const range = CellRange.parse('A1:A3');
    FilterEngine.apply(sheet, range, filterSpec(new Map([[0, { allowedValues: new Set(['Open']) }]])), INV);
    expect(sheet.isRowHidden(2)).toBe(true);

    sheet.activeFilter = null;
    sheet.activeFilterSpec = null;
    history.execute(DataEditCommands.applySortAndFilter(sheet.name, range, null, null, INV));

    expect(sheet.isRowHidden(2)).toBe(false);
  });
});
