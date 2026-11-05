import { describe, expect, it } from 'vitest';
import { table, type ColumnFilter, type FilterSpec, type SortSpec } from '../../src/data/types.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { writeOds } from '../../src/serialization/index.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { readEntry, roundTrip } from './helpers.ts';

/**
 * Round-trips the ODF `table:database-range` that carries a worksheet's filter region, filter
 * conditions, and sort order. A filter + sort must survive save → reopen, and the rows the filter
 * hides must come back hidden.
 *
 * The C# suite applies the filter with `FilterEngine.Apply` before saving; the TS engine has no
 * FilterEngine yet, so these tests hide the same rows by hand (what Apply would have done) and the
 * reader relies on the persisted `table:visibility` rather than re-running the filter.
 */
function filterSpec(columns: Record<number, ColumnFilter>, hasHeader = true): FilterSpec {
  return { columns: new Map(Object.entries(columns).map(([k, v]) => [Number(k), v])), hasHeader };
}

function sortSpec(levels: Array<[column: number, ascending?: boolean]>, hasHeader = true): SortSpec {
  return { levels: levels.map(([column, ascending = true]) => ({ column, ascending })), hasHeader };
}

describe('OdsDatabaseRange', () => {
  it('header column and banded rows flags round trip', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.text('Site'));
    sheet.setValue(1, 0, CellValue.text('All Sites'));
    sheet.addTable(
      table('T', CellRange.parse('A1:B3'), { hasHeader: true, hasHeaderColumn: true, hasBandedRows: true }),
    );

    const tables = roundTrip(workbook).findSheet('Sheet1')!.tables;
    expect(tables).toHaveLength(1);
    expect(tables[0]!.hasHeaderColumn).toBe(true);
    expect(tables[0]!.hasBandedRows).toBe(true);
  });

  it('a table without the flags writes no cellestra attributes', () => {
    // Off is the default and stays invisible in the file: pure vocabulary ODF.
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.text('Site'));
    sheet.addTable(table('T', CellRange.parse('A1:B3'), { hasHeader: true }));

    expect(readEntry(writeOds(workbook), 'content.xml')).not.toContain('urn:cellestra');

    const tables = roundTrip(workbook).findSheet('Sheet1')!.tables;
    expect(tables).toHaveLength(1);
    expect(tables[0]!.hasHeaderColumn).toBe(false);
    expect(tables[0]!.hasBandedRows).toBe(false);
  });

  function buildWithSheetFilter(): Workbook {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;

    // Header + four data rows.
    sheet.setValue(0, 0, CellValue.text('Name'));
    sheet.setValue(0, 1, CellValue.text('Qty'));
    sheet.setValue(1, 0, CellValue.text('apple'));
    sheet.setValue(1, 1, CellValue.number(3));
    sheet.setValue(2, 0, CellValue.text('banana'));
    sheet.setValue(2, 1, CellValue.number(8));
    sheet.setValue(3, 0, CellValue.text('cherry'));
    sheet.setValue(3, 1, CellValue.number(1));
    sheet.setValue(4, 0, CellValue.text('apple'));
    sheet.setValue(4, 1, CellValue.number(9));

    const range = CellRange.parse('A1:B5');

    // Column 0: a value checklist (apple only); Column 1: a relational condition (> 2).
    const spec = filterSpec({ 0: { allowedValues: new Set(['apple']) }, 1: { condition: '>2' } });
    const sort = sortSpec([[1, false]]);

    // What FilterEngine.Apply does in C#: apple AND qty>2 hides rows 2 (banana) and 3 (cherry).
    sheet.getOrCreateRow(2).hidden = true;
    sheet.getOrCreateRow(3).hidden = true;
    sheet.activeFilter = range;
    sheet.activeFilterSpec = spec;
    sheet.activeSort = sort;
    return workbook;
  }

  it('sheet filter region, conditions, and sort survive', () => {
    const sheet = roundTrip(buildWithSheetFilter()).findSheet('Sheet1')!;

    expect(sheet.activeFilter!.equals(CellRange.parse('A1:B5'))).toBe(true);

    const spec = sheet.activeFilterSpec;
    expect(spec).not.toBeNull();
    expect(spec!.hasHeader).toBe(true);
    expect(spec!.columns.size).toBe(2);
    expect([...spec!.columns.get(0)!.allowedValues!].sort()).toEqual(['apple']);
    expect(spec!.columns.get(1)!.condition).toBe('>2');

    const sort = sheet.activeSort;
    expect(sort).not.toBeNull();
    expect(sort!.levels).toHaveLength(1);
    expect(sort!.levels[0]!.column).toBe(1);
    expect(sort!.levels[0]!.ascending).toBe(false);
  });

  it('rows hidden by the filter stay hidden on open', () => {
    const sheet = roundTrip(buildWithSheetFilter()).findSheet('Sheet1')!;

    // apple AND qty>2 → rows 1 and 4 visible; 2 (banana) and 3 (cherry) hidden.
    expect(sheet.isRowHidden(1)).toBe(false);
    expect(sheet.isRowHidden(2)).toBe(true);
    expect(sheet.isRowHidden(3)).toBe(true);
    expect(sheet.isRowHidden(4)).toBe(false);
  });

  it('table filter and sort survive', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.text('City'));
    sheet.setValue(1, 0, CellValue.text('Oslo'));
    sheet.setValue(2, 0, CellValue.text('Lima'));

    const range = CellRange.parse('A1:A3');
    const filter = filterSpec({ 0: { allowedValues: new Set(['Oslo', 'Lima']) } });
    sheet.addTable(table('Cities', range, { hasHeader: true, filter, sort: sortSpec([[0]]) }));

    const tables = roundTrip(workbook).findSheet('Sheet1')!.tables;
    expect(tables).toHaveLength(1);
    const t = tables[0]!;
    expect(t.name).toBe('Cities');
    expect(t.range.equals(range)).toBe(true);
    expect(t.filter).toBeTruthy();
    expect(t.filter!.columns.get(0)!.allowedValues!.size).toBe(2);
    expect(t.sort).toBeTruthy();
    expect(t.sort!.levels).toHaveLength(1);
    expect(t.sort!.levels[0]!.column).toBe(0);
  });

  it('multi-value checklist round trips via filter-or', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.text('Color'));
    sheet.setValue(1, 0, CellValue.text('red'));
    sheet.setValue(2, 0, CellValue.text('green'));
    sheet.setValue(3, 0, CellValue.text('blue'));

    const range = CellRange.parse('A1:A4');
    sheet.getOrCreateRow(2).hidden = true; // FilterEngine.Apply: green is not in the checklist
    sheet.activeFilter = range;
    sheet.activeFilterSpec = filterSpec({ 0: { allowedValues: new Set(['red', 'blue']) } });

    const reopened = roundTrip(workbook).findSheet('Sheet1')!;
    expect([...reopened.activeFilterSpec!.columns.get(0)!.allowedValues!].sort()).toEqual(['blue', 'red']);
  });

  it('multi-level sort and multi-column filter survive', () => {
    // The chained pipeline a Sort & Filter dialog produces: sort by two columns, filter on two.
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.text('Region'));
    sheet.setValue(0, 1, CellValue.text('Status'));
    sheet.setValue(1, 0, CellValue.text('West'));
    sheet.setValue(1, 1, CellValue.text('Open'));
    sheet.setValue(2, 0, CellValue.text('East'));
    sheet.setValue(2, 1, CellValue.text('Done'));

    sheet.activeFilter = CellRange.parse('A1:B3');
    sheet.activeFilterSpec = filterSpec({
      0: { allowedValues: new Set(['West', 'East']) },
      1: { condition: 'Open' },
    });
    sheet.activeSort = sortSpec([
      [0, true],
      [1, false],
    ]);

    const reopened = roundTrip(workbook).findSheet('Sheet1')!;

    const sort = reopened.activeSort;
    expect(sort).not.toBeNull();
    expect(sort!.levels.map((l) => l.column)).toEqual([0, 1]); // both levels, in priority order
    expect(sort!.levels[0]!.ascending).toBe(true);
    expect(sort!.levels[1]!.ascending).toBe(false);
    expect(reopened.activeFilterSpec!.columns.size).toBe(2);
  });

  it('no filter emits no database range', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.setValue(0, 0, CellValue.number(1));

    const reopened = roundTrip(workbook).findSheet('Sheet1')!;
    expect(reopened.activeFilter).toBeNull();
    expect(reopened.activeFilterSpec).toBeNull();
    expect(reopened.tables).toHaveLength(0);
  });
});
