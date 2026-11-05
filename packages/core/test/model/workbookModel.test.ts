import { createChart } from '@cellestra/charts';
import { describe, expect, it } from 'vitest';
import { Defaults } from '../../src/defaults.ts';
import { CellAddress } from '../../src/model/cellAddress.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { MergedRegion } from '../../src/model/mergedRegion.ts';
import { SheetName } from '../../src/model/sheetName.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const textAt = (sheet: Worksheet, row: number, column: number): string => {
  const v = sheet.getValue(row, column);
  if (v.kind !== 'text') throw new Error(`Expected text, got ${v.kind}.`);
  return v.value;
};

const sameRange = (actual: CellRange | null, expected: CellRange): void => {
  expect(actual?.equals(expected), `${String(actual)} vs ${expected.toString()}`).toBe(true);
};

describe('Workbook model', () => {
  it('createNew has one sheet named Sheet1', () => {
    const workbook = Workbook.createNew();
    expect(workbook.sheetCount).toBe(1);
    expect(workbook.activeSheet.name).toBe('Sheet1');
  });

  it('addSheet auto-names the next unused number', () => {
    const workbook = Workbook.createNew();
    expect(workbook.addSheet().name).toBe('Sheet2');
    expect(workbook.addSheet().name).toBe('Sheet3');
  });

  it('addSheet skips taken numbers', () => {
    const workbook = Workbook.createNew();
    workbook.addSheet('Sheet2');
    expect(workbook.addSheet().name).toBe('Sheet3');
  });

  it('duplicateSheet inserts an independent copy right after the original', () => {
    const workbook = Workbook.createNew();
    const source = workbook.activeSheet; // Sheet1
    source.setValue(0, 0, CellValue.text('hello'));
    source.frozenRows = 1;
    workbook.addSheet('After'); // a sheet after Sheet1, to prove the copy lands between them

    const copy = workbook.duplicateSheet(source);

    expect(copy.name).toBe('Sheet1 copy');
    expect(workbook.indexOf(copy)).toBe(1); // immediately after the original
    expect(textAt(copy, 0, 0)).toBe('hello'); // content copied
    expect(copy.frozenRows).toBe(1); // layout copied

    // The copy is fully detached: editing it leaves the original untouched.
    copy.setValue(0, 0, CellValue.text('changed'));
    expect(textAt(source, 0, 0)).toBe('hello');
  });

  it('duplicateSheet copies a chart sheet with its own chart', () => {
    const workbook = Workbook.createNew();
    const source = workbook.addChartSheet(createChart({ title: 'Sales' }));

    const copy = workbook.duplicateSheet(source);

    expect(copy.isChartSheet).toBe(true);
    expect(copy.chart!.title).toBe('Sales');
    expect(copy.chart).not.toBe(source.chart); // a deep copy, not a shared reference
  });

  it('duplicateSheet renames copied tables so they stay unique', () => {
    const workbook = Workbook.createNew();
    const source = workbook.activeSheet;
    source.setValue(0, 0, CellValue.text('H'));
    source.setValue(1, 0, CellValue.number(1));
    source.addTable({ name: 'Sales', range: CellRange.parse('A1:A2'), hasHeader: true });

    const copy = workbook.duplicateSheet(source);

    expect(source.tables[0]!.name).toBe('Sales'); // original keeps its name
    expect(copy.tables[0]!.name).toBe('Sales_2'); // the copy's table is renamed unique
  });

  it('duplicateSheet name avoids collisions', () => {
    const workbook = Workbook.createNew();
    const source = workbook.activeSheet;
    expect(workbook.duplicateSheet(source).name).toBe('Sheet1 copy');
    expect(workbook.duplicateSheet(source).name).toBe('Sheet1 copy 2');
  });

  it('cannot remove the last sheet', () => {
    const workbook = Workbook.createNew();
    expect(() => workbook.removeSheet(workbook.activeSheet)).toThrow();
  });

  it.each(['', '   ', 'Has/Slash', 'Bad:Colon', 'Q?', 'Star*', 'Bracket[1]', "'quoted'"])(
    'addSheet rejects invalid name %j',
    (name) => {
      expect(() => Workbook.createNew().addSheet(name)).toThrow(RangeError);
    },
  );

  it('addSheet rejects names over the length cap', () => {
    expect(() => Workbook.createNew().addSheet('x'.repeat(SheetName.maxLength + 1))).toThrow(RangeError);
  });

  it('names with spaces are allowed', () => {
    expect(Workbook.createNew().addSheet('My Sheet').name).toBe('My Sheet');
  });

  it('renameSheet enforces uniqueness and rules', () => {
    const workbook = Workbook.createNew();
    const sheet1 = workbook.activeSheet;
    workbook.addSheet('Data');

    expect(() => workbook.renameSheet(sheet1, 'Data')).toThrow(/already exists/); // duplicate
    expect(() => workbook.renameSheet(sheet1, 'no/slash')).toThrow(RangeError); // invalid

    workbook.renameSheet(sheet1, 'Summary');
    expect(sheet1.name).toBe('Summary');
    workbook.renameSheet(sheet1, 'Summary'); // renaming to its own name is a no-op
    expect(sheet1.name).toBe('Summary');
  });

  it('a worksheet stores only non-blank cells', () => {
    const sheet = new Worksheet('S');
    sheet.setValue(0, 0, CellValue.number(42));
    expect(sheet.cellCount).toBe(1);
    sheet.setValue(0, 0, CellValue.empty);
    expect(sheet.cellCount).toBe(0);
  });

  it('getValue of a blank cell is empty', () => {
    expect(new Worksheet('S').getValue(10, 10).kind).toBe('empty');
  });

  it('setFormula keeps the cached value', () => {
    const sheet = new Worksheet('S');
    sheet.setFormula(0, 0, 'SUM(A2:A3)', CellValue.number(7));

    const cell = sheet.find(0, 0)!;
    expect(cell.isFormula).toBe(true);
    expect(cell.formula).toBe('SUM(A2:A3)');
    expect(cell.value).toEqual(CellValue.number(7));
  });

  it('merged regions reject overlap', () => {
    const sheet = new Worksheet('S');
    sheet.addMerge(new MergedRegion(CellRange.parse('A1:B2')));
    expect(() => sheet.addMerge(new MergedRegion(CellRange.parse('B2:C3')))).toThrow();
  });

  it('selectionRangeAt returns the whole merge for any covered cell', () => {
    const sheet = new Worksheet('S');
    sheet.addMerge(new MergedRegion(CellRange.parse('B2:D4')));

    sameRange(sheet.selectionRangeAt(1, 1), CellRange.parse('B2:D4')); // anchor B2
    sameRange(sheet.selectionRangeAt(3, 3), CellRange.parse('B2:D4')); // covered D4
    sameRange(sheet.selectionRangeAt(2, 2), CellRange.parse('B2:D4')); // covered C3
    // A cell outside the merge is just itself.
    sameRange(sheet.selectionRangeAt(0, 0), new CellRange(new CellAddress(0, 0)));
  });

  it('stepOverMerges skips past a merge when leaving it', () => {
    const sheet = new Worksheet('S');
    sheet.addMerge(new MergedRegion(CellRange.parse('B2:D4'))); // rows 1..3, cols 1..3

    expect(sheet.stepOverMerges(1, 1, 1, 0)).toEqual([4, 1]); // down: row below the bottom edge
    expect(sheet.stepOverMerges(1, 1, 0, 1)).toEqual([1, 4]); // right: column past the right edge
    expect(sheet.stepOverMerges(2, 2, -1, 0)).toEqual([0, 2]); // up: keeps the cell's own column
    expect(sheet.stepOverMerges(2, 2, 0, -1)).toEqual([2, 0]); // left: keeps the cell's own row
  });

  it('stepOverMerges snaps to the anchor when entering a merge', () => {
    const sheet = new Worksheet('S');
    sheet.addMerge(new MergedRegion(CellRange.parse('B2:D4'))); // anchor B2 = (1,1)

    expect(sheet.stepOverMerges(0, 2, 1, 0)).toEqual([1, 1]);
    expect(sheet.stepOverMerges(3, 0, 0, 1)).toEqual([1, 1]);
  });

  it('stepOverMerges clamps at sheet edges', () => {
    expect(new Worksheet('S').stepOverMerges(0, 0, -1, -1)).toEqual([0, 0]);
  });

  it('getUsedRange covers content and sizing', () => {
    const sheet = new Worksheet('S');
    sheet.setValue(2, 3, CellValue.text('x'));
    sheet.getOrCreateColumn(5).width = 100;

    const used = sheet.getUsedRange();

    expect(used).not.toBeNull();
    expect([used!.top, used!.bottom, used!.left, used!.right]).toEqual([2, 2, 3, 5]);
  });

  it('named ranges resolve sheet scope over global', () => {
    const workbook = Workbook.createNew();
    workbook.addNamedRange({ name: 'Rate', target: CellRange.parse('A1') });
    workbook.addNamedRange({ name: 'Rate', target: CellRange.parse('B2'), scope: 'Sheet1' });

    sameRange(workbook.findNamedRange('Rate', 'Sheet1')!.target, CellRange.parse('B2'));
    sameRange(workbook.findNamedRange('Rate')!.target, CellRange.parse('A1'));
  });

  it('defaults match PRD §8.8', () => {
    expect(Defaults.fontFamily).toBe('Helvetica Neue');
    expect(Defaults.fontSizePoints).toBe(10);
    expect(Defaults.columnWidth).toBe(64);
    expect(Defaults.rowHeight).toBe(20);
  });
});
