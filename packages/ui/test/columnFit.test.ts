import { CellValue, NumberFormatter, OdsWriter, Workbook } from '@cellestra/core';
import { describe, expect, it } from 'vitest';
import { CELL_PAD_X, type TextMeasure } from '../src/grid/cellText.ts';
import { CSV_FIT_MAX_WIDTH, CSV_FIT_ROWS, fitCsvColumns, fittedColumnWidths } from '../src/grid/columnFit.ts';
import { WorkbookSession } from '../src/session/workbookSession.ts';

/** A fake canvas: every character is 10px wide. */
const measure = (): TextMeasure => ({ font: '', measureText: (t) => ({ width: t.length * 10 }) });
const pad = 2 * CELL_PAD_X;

function sheetWith(rows: readonly (readonly string[])[]) {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  rows.forEach((cells, r) => cells.forEach((t, c) => t && sheet.setValue(r, c, CellValue.text(t))));
  return { workbook, sheet };
}

const display = (locale: string) => (v: CellValue, s: { numberFormat: { code: string } }) =>
  NumberFormatter.format(v, s.numberFormat.code, locale);

describe('fittedColumnWidths', () => {
  it('fits each used column to its widest value, leaving empty columns out', () => {
    const { workbook, sheet } = sheetWith([
      ['ab', '', 'abcdef'],
      ['abcd', '', 'a'],
    ]);
    const widths = fittedColumnWidths(measure(), workbook, sheet, display('en-US'));
    expect([...widths].sort()).toEqual([
      [0, 40 + pad],
      [2, 60 + pad],
    ]);
  });

  it('honours a column filter, a row limit and a maximum width', () => {
    const { workbook, sheet } = sheetWith([
      ['a', 'b'],
      ['x'.repeat(100), 'bbbb'],
    ]);
    const widths = fittedColumnWidths(measure(), workbook, sheet, display('en-US'), {
      columns: (c) => c === 0,
      maxWidth: 300,
    });
    expect([...widths]).toEqual([[0, 300]]);
    const firstRow = fittedColumnWidths(measure(), workbook, sheet, display('en-US'), { rowLimit: 1 });
    expect(firstRow.get(0)).toBe(10 + pad);
  });
});

describe('fitCsvColumns', () => {
  it('caps a long column and measures only the first rows', () => {
    const { workbook, sheet } = sheetWith([['short', 'y'.repeat(200)]]);
    sheet.setValue(CSV_FIT_ROWS, 0, CellValue.text('z'.repeat(30)));
    fitCsvColumns(measure(), workbook, sheet);
    expect(sheet.columns.get(0)?.width).toBe(50 + pad);
    expect(sheet.columns.get(1)?.width).toBe(CSV_FIT_MAX_WIDTH);
  });
});

describe('opening a CSV', () => {
  const csv = new TextEncoder().encode('id,description\n1,a longer description\n');
  const ref = { store: 'browser' as const, id: 'd', name: 'data.csv', version: null };

  it('fits its columns as part of loading: not modified, nothing to undo', () => {
    const session = new WorkbookSession();
    session.textMeasure = measure;
    session.load(csv, ref);
    const sheet = session.activeSheet.peek();
    expect(sheet.columns.get(0)?.width).toBe(20 + pad);
    expect(sheet.columns.get(1)?.width).toBe(200 + pad);
    expect(session.modified.peek()).toBe(false);
    expect(session.undoLabel).toBeNull();
  });

  it('fits them again when the CSV is reloaded', () => {
    const session = new WorkbookSession();
    session.textMeasure = measure;
    session.load(csv, ref);
    session.reload(new TextEncoder().encode('id\n123456\n'), ref);
    expect(session.activeSheet.peek().columns.get(0)?.width).toBe(60 + pad);
  });

  it('leaves an .ods file’s widths as saved', () => {
    const session = new WorkbookSession();
    session.textMeasure = measure;
    const { workbook } = sheetWith([['a long value here']]);
    session.load(OdsWriter.write(workbook, null), { ...ref, name: 'book.ods' });
    expect(session.activeSheet.peek().columns.get(0)?.width ?? 64).toBe(64);
  });
});
