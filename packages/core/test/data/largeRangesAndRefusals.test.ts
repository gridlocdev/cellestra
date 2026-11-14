import { describe, expect, it } from 'vitest';
import { EdgeNavigator } from '../../src/data/edgeNavigator.ts';
import { filterSpec } from '../../src/data/filterEngine.ts';
import { dataValidation } from '../../src/data/types.ts';
import { ValidationPreflight } from '../../src/data/validationPreflight.ts';
import { DataEditCommands } from '../../src/editing/dataEditCommands.ts';
import { EditRefusedError } from '../../src/editing/editCommand.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { MergedRegion } from '../../src/model/mergedRegion.ts';
import { SheetDimensions } from '../../src/model/sheetDimensions.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { NumberFormats } from '../../src/styling/cellStyle.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const R = (a1: string): CellRange => CellRange.parse(a1);

function setup() {
  const workbook = Workbook.createNew();
  return { workbook, history: new EditHistory(workbook), sheet: workbook.activeSheet };
}

describe('sort over merged cells', () => {
  it('is refused before anything moves, and leaves no undo step', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('H'));
    [3, 1, 2].forEach((v, i) => sheet.setValue(i + 1, 0, CellValue.number(v)));
    sheet.setValue(1, 1, CellValue.text('merged'));
    sheet.addMerge(new MergedRegion(R('B2:C2')));

    expect(() => history.execute(DataEditCommands.sortByColumn(sheet.name, R('A1:C4'), 0, true))).toThrow(
      EditRefusedError,
    );
    expect([1, 2, 3].map((r) => CellValue.asNumber(sheet.getValue(r, 0)))).toEqual([3, 1, 2]);
    expect(history.canUndo).toBe(false);
  });

  it('a merged header row does not block the sort', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('H'));
    sheet.addMerge(new MergedRegion(R('A1:B1')));
    [3, 1, 2].forEach((v, i) => sheet.setValue(i + 1, 0, CellValue.number(v)));
    history.execute(DataEditCommands.sortByColumn(sheet.name, R('A1:B4'), 0, true));
    expect([1, 2, 3].map((r) => CellValue.asNumber(sheet.getValue(r, 0)))).toEqual([1, 2, 3]);
  });

  it('a refused step in Sort & Filter takes back the filter it already applied', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('H'));
    [3, 1, 2].forEach((v, i) => sheet.setValue(i + 1, 0, CellValue.number(v)));
    sheet.addMerge(new MergedRegion(R('B2:B3')));
    const command = DataEditCommands.applySortAndFilter(
      sheet.name,
      R('A1:B4'),
      { levels: [{ column: 0, ascending: true }], hasHeader: true },
      filterSpec(new Map([[0, { condition: '<>1' }]])),
      'en-US',
    );
    expect(() => history.execute(command)).toThrow(EditRefusedError);
    expect(sheet.isRowHidden(2)).toBe(false);
    expect(sheet.activeFilter).toBeNull();
  });
});

describe('filtering whole columns', () => {
  it('stops at the data, creates rows only where visibility changes, and undoes exactly', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('H'));
    [1, -1, 2, -2].forEach((v, i) => sheet.setValue(i + 1, 0, CellValue.number(v)));
    sheet.getOrCreateRow(10).manuallyHidden = true; // hidden by hand, past the data

    history.execute(
      DataEditCommands.applyFilter(
        sheet.name,
        R('A1:A1048576'),
        filterSpec(new Map([[0, { condition: '>0' }]])),
        'en-US',
      ),
    );
    expect([1, 2, 3, 4].map((r) => sheet.isRowHidden(r))).toEqual([false, true, false, true]);
    expect(sheet.rows.size).toBe(3); // the two filtered rows and the hand-hidden one
    expect(sheet.isRowHidden(10)).toBe(true);
    expect(sheet.isRowHidden(500000)).toBe(false);

    history.undo();
    expect([1, 2, 3, 4].map((r) => sheet.isRowHidden(r))).toEqual([false, false, false, false]);
    expect(sheet.isRowHidden(10)).toBe(true);
  });

  it('clearing a whole-column filter unhides its rows and undo hides them again', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('H'));
    sheet.setValue(1, 0, CellValue.number(1));
    sheet.getOrCreateRow(1).filtered = true;
    history.execute(DataEditCommands.clearFilter(sheet.name, R('A1:A1048576')));
    expect(sheet.isRowHidden(1)).toBe(false);
    history.undo();
    expect(sheet.isRowHidden(1)).toBe(true);
  });
});

describe('text to columns', () => {
  it('reads only the used rows of a whole column', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('a,b,c'));
    const command = DataEditCommands.textToColumns(sheet, R('A1:A1048576'), ',');
    history.execute(command);
    expect(sheet.getValue(0, 2)).toEqual(CellValue.text('c'));
    history.undo();
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('a,b,c'));
    expect(sheet.getValue(0, 1).kind).toBe('empty');
  });

  it('reads fields as typed input: numbers, booleans and dates (with a date format)', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('1,2.5,TRUE,2024-03-01,=A1,x'));
    history.execute(DataEditCommands.textToColumns(sheet, R('A1:A1'), ',', 'en-US'));
    expect(sheet.getValue(0, 0)).toEqual(CellValue.number(1));
    expect(sheet.getValue(0, 1)).toEqual(CellValue.number(2.5));
    expect(sheet.getValue(0, 2)).toEqual(CellValue.boolean(true));
    expect(sheet.getValue(0, 3).kind).toBe('number');
    const dateStyle = workbook.styles.get(sheet.effectiveStyleId(0, 3));
    expect(NumberFormats.isGeneral(dateStyle.numberFormat)).toBe(false);
    expect(sheet.getValue(0, 4)).toEqual(CellValue.text('=A1')); // never a formula
    expect(sheet.find(0, 4)?.formula).toBeNull();
    expect(sheet.getValue(0, 5)).toEqual(CellValue.text('x'));

    history.undo();
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('1,2.5,TRUE,2024-03-01,=A1,x'));
    expect(sheet.find(0, 3)).toBeUndefined();
  });
});

describe('edge navigation on a sparse sheet', () => {
  it('jumps from an empty column without walking to the sheet edge cell by cell', () => {
    const { sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('x'));
    expect(EdgeNavigator.findEdge(sheet, 5, 3, 1, 0)).toEqual([SheetDimensions.maxRows - 1, 3]);
    expect(EdgeNavigator.findEdge(sheet, 900000, 3, -1, 0)).toEqual([0, 3]);
  });

  it('still finds the next filled cell far away, in either direction', () => {
    const { sheet } = setup();
    sheet.setValue(800000, 3, CellValue.number(1));
    sheet.setValue(10, 3, CellValue.number(2));
    expect(EdgeNavigator.findEdge(sheet, 20, 3, 1, 0)).toEqual([800000, 3]);
    expect(EdgeNavigator.findEdge(sheet, 700000, 3, -1, 0)).toEqual([10, 3]);
    expect(EdgeNavigator.findEdge(sheet, 10, 5000, 0, -1)).toEqual([10, 3]);
  });
});

describe('validation pre-flight', () => {
  it('reads a list source once for a whole batch of writes', () => {
    const { sheet } = setup();
    sheet.addValidation(dataValidation('list', R('A1:A100'), { listSource: R('C1:C3') }));
    let reads = 0;
    const resolve = (): CellValue[] => {
      reads++;
      return [CellValue.text('a'), CellValue.text('b')];
    };
    const incoming = Array.from({ length: 100 }, (_, row) => ({
      row,
      column: 0,
      value: CellValue.text(row % 2 ? 'a' : 'z'),
      isFormula: false,
    }));
    const bad = ValidationPreflight.findStopViolations(sheet, incoming, 'en-US', resolve);
    expect(bad).toHaveLength(50);
    expect(reads).toBe(1);
  });
});
