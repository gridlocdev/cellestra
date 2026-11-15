import { describe, expect, it } from 'vitest';
import { dataValidation, table } from '../../src/data/types.ts';
import { DuplicateRemover, TextToColumns } from '../../src/data/dataTools.ts';
import { listItems, Validator } from '../../src/data/dataValidation.ts';
import { filterSpec, FilterEngine } from '../../src/data/filterEngine.ts';
import { sortSpec, SortEngine } from '../../src/data/sortEngine.ts';
import { TableAutoExpand, tableHeaderLabels } from '../../src/data/tables.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { numberAt, textAt } from './helpers.ts';

const INV = 'en-US';

function sheetWith(...cells: [number, number, number | string][]): Worksheet {
  const sheet = new Worksheet('S');
  for (const [row, col, value] of cells) {
    sheet.setValue(row, col, typeof value === 'number' ? CellValue.number(value) : CellValue.text(value));
  }
  return sheet;
}

describe('DataEngine', () => {
  it('sort orders data rows keeping header', () => {
    const sheet = sheetWith(
      [0, 0, 'Name'],
      [0, 1, 'Score'],
      [1, 0, 'Bev'],
      [1, 1, 87],
      [2, 0, 'Ada'],
      [2, 1, 99],
      [3, 0, 'Cy'],
      [3, 1, 50],
    );
    SortEngine.sort(sheet, CellRange.parse('A1:B4'), sortSpec([{ column: 0, ascending: true }]));
    expect(textAt(sheet, 0, 0)).toBe('Name');
    expect(textAt(sheet, 1, 0)).toBe('Ada');
    expect(textAt(sheet, 2, 0)).toBe('Bev');
    expect(textAt(sheet, 3, 0)).toBe('Cy');
  });

  it('multi-level sort is stable', () => {
    const sheet = sheetWith(
      [0, 0, 'Group'],
      [0, 1, 'N'],
      [1, 0, 'A'],
      [1, 1, 2],
      [2, 0, 'A'],
      [2, 1, 1],
      [3, 0, 'B'],
      [3, 1, 1],
    );
    SortEngine.sort(
      sheet,
      CellRange.parse('A1:B4'),
      sortSpec([
        { column: 0, ascending: true },
        { column: 1, ascending: true },
      ]),
    );
    expect(numberAt(sheet, 1, 1)).toBe(1);
    expect(numberAt(sheet, 2, 1)).toBe(2);
    expect(textAt(sheet, 3, 0)).toBe('B');
  });

  it('filter hides non-matching rows', () => {
    const sheet = sheetWith([0, 0, 'Score'], [1, 0, 10], [2, 0, 50], [3, 0, 90]);
    const spec = filterSpec(new Map([[0, { condition: '>40' }]]));
    FilterEngine.apply(sheet, CellRange.parse('A1:A4'), spec, INV);
    expect(sheet.isRowHidden(1)).toBe(true);
    expect(sheet.isRowHidden(2)).toBe(false);
    expect(sheet.isRowHidden(3)).toBe(false);
    FilterEngine.clear(sheet, CellRange.parse('A1:A4'));
    expect(sheet.isRowHidden(1)).toBe(false);
  });

  it('sort reapplies the active filter so visibility follows the data', () => {
    const sheet = sheetWith([0, 0, 'Score'], [1, 0, 90], [2, 0, 10], [3, 0, 50]);
    const range = CellRange.parse('A1:A4');
    const spec = filterSpec(new Map([[0, { condition: '>40' }]]));
    FilterEngine.apply(sheet, range, spec, INV);
    sheet.activeFilter = range;
    sheet.activeFilterSpec = spec;
    expect(sheet.isRowHidden(1)).toBe(false);
    expect(sheet.isRowHidden(2)).toBe(true);
    expect(sheet.isRowHidden(3)).toBe(false);

    SortEngine.sort(sheet, range, sortSpec([{ column: 0, ascending: true }]), INV);

    expect(numberAt(sheet, 1, 0)).toBe(10);
    expect(numberAt(sheet, 2, 0)).toBe(50);
    expect(numberAt(sheet, 3, 0)).toBe(90);
    expect(sheet.isRowHidden(1)).toBe(true);
    expect(sheet.isRowHidden(2)).toBe(false);
    expect(sheet.isRowHidden(3)).toBe(false);
  });

  it('remove duplicates keeps first occurrence', () => {
    const sheet = sheetWith([0, 0, 'x'], [1, 0, 'a'], [2, 0, 'b'], [3, 0, 'a'], [4, 0, 'b'], [5, 0, 'c']);
    const removed = DuplicateRemover.remove(sheet, CellRange.parse('A1:A6'), [0], true);
    expect(removed).toBe(2);
    expect(textAt(sheet, 1, 0)).toBe('a');
    expect(textAt(sheet, 2, 0)).toBe('b');
    expect(textAt(sheet, 3, 0)).toBe('c');
    expect(sheet.getValue(4, 0).kind).toBe('empty');
  });

  it('text to columns splits on delimiter', () => {
    const sheet = sheetWith([0, 0, 'a,b,c'], [1, 0, 'd,e']);
    TextToColumns.split(sheet, CellRange.parse('A1:A2'), ',');
    expect(textAt(sheet, 0, 0)).toBe('a');
    expect(textAt(sheet, 0, 1)).toBe('b');
    expect(textAt(sheet, 0, 2)).toBe('c');
    expect(textAt(sheet, 1, 1)).toBe('e');
  });

  it('table auto-expands below and right', () => {
    const t = table('Data', CellRange.parse('A1:B3'));
    expect(TableAutoExpand.expandFor(t, 3, 1)?.equals(CellRange.parse('A1:B4'))).toBe(true);
    expect(TableAutoExpand.expandFor(t, 0, 2)?.equals(CellRange.parse('A1:C3'))).toBe(true);
    expect(TableAutoExpand.expandFor(t, 10, 10)).toBeNull();
  });

  it.each([
    [5, true],
    [15, false],
  ])('validation number between: %d -> %s', (input, valid) => {
    const rule = dataValidation('wholeNumber', CellRange.parse('A1'), {
      operator: 'between',
      operand1: 1,
      operand2: 10,
    });
    expect(Validator.isValid(rule, CellValue.number(input), INV)).toBe(valid);
  });

  it('validation list accepts only listed values', () => {
    const rule = dataValidation('list', CellRange.parse('A1'), { listValues: ['Red', 'Green', 'Blue'] });
    expect(Validator.isValid(rule, CellValue.text('Green'), INV)).toBe(true);
    expect(Validator.isValid(rule, CellValue.text('Purple'), INV)).toBe(false);
  });

  it('list items skip blank and whitespace cells, trim, and drop repeats ignoring case', () => {
    const source = CellRange.parse('B1:B6');
    const rule = dataValidation('list', CellRange.parse('A1'), { listSource: source });
    const cells = [
      CellValue.text(' Red '),
      CellValue.text(''),
      CellValue.text('   '),
      CellValue.empty,
      CellValue.text('red'),
      CellValue.number(2),
    ];
    expect(listItems(rule, () => cells)).toEqual(['Red', '2']);
    const typed = dataValidation('list', CellRange.parse('A1'), { listValues: ['a', ' ', 'A', ' b'] });
    expect(listItems(typed, null)).toEqual(['a', 'b']);
  });

  it('list validation matches trimmed text and accepts a blank entry', () => {
    const rule = dataValidation('list', CellRange.parse('A1'), { listSource: CellRange.parse('B1:B2') });
    const resolve = () => [CellValue.text(' Red '), CellValue.text('')];
    expect(Validator.isValid(rule, CellValue.text('red'), INV, resolve)).toBe(true);
    expect(Validator.isValid(rule, CellValue.text('  '), INV, resolve)).toBe(true);
    expect(Validator.isValid(rule, CellValue.text('Blue'), INV, resolve)).toBe(false);
  });
});

describe('table header labels', () => {
  it('name blank headers by column and repeated headers uniquely', () => {
    const sheet = new Worksheet('S');
    ['Sales', ' ', 'Sales', 'sales'].forEach((h, c) => sheet.setValue(0, c, CellValue.text(h)));
    const t = table('T', CellRange.parse('A1:E3'));
    expect(tableHeaderLabels(t, sheet)).toEqual(['Sales', 'Column2', 'Sales (2)', 'sales (3)', 'Column5']);
  });
});
