import { describe, expect, it } from 'vitest';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { ImportCsvCommand } from '../../src/editing/importCsvCommand.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CsvOptions } from '../../src/serialization/csv/csvOptions.ts';
import { CellValue } from '../../src/values/cellValue.ts';

function setup() {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  sheet.setValue(0, 0, CellValue.text('old'));
  sheet.setValue(9, 9, CellValue.number(7));
  sheet.getOrCreateColumn(2).width = 5;
  sheet.frozenRows = 2;
  return { workbook, sheet, history: new EditHistory(workbook) };
}

describe('ImportCsvCommand', () => {
  it('replaces the sheet with the CSV, typed per the import rules', () => {
    const { workbook, sheet, history } = setup();
    history.execute(
      new ImportCsvCommand(sheet.name, '﻿name,amount\n"a, b",1.5\n=1+1,007\n', CsvOptions.comma),
    );
    expect(workbook.activeSheet).toBe(sheet);
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('name'));
    expect(sheet.getValue(1, 0)).toEqual(CellValue.text('a, b'));
    expect(sheet.getValue(1, 1)).toEqual(CellValue.number(1.5));
    expect(sheet.getValue(2, 0)).toEqual(CellValue.text('=1+1'));
    expect(sheet.getValue(2, 1)).toEqual(CellValue.text('007'));
    expect(sheet.getValue(9, 9)).toEqual(CellValue.empty);
    expect(sheet.frozenRows).toBe(0);
    expect(sheet.columns.get(2)).toBeUndefined();
  });

  it('brings the old sheet back in one undo, and redoes the import', () => {
    const { sheet, history } = setup();
    history.execute(new ImportCsvCommand(sheet.name, 'x\ty\n', CsvOptions.tab));
    expect(sheet.getValue(0, 1)).toEqual(CellValue.text('y'));
    history.undo();
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('old'));
    expect(sheet.getValue(0, 1)).toEqual(CellValue.empty);
    expect(sheet.getValue(9, 9)).toEqual(CellValue.number(7));
    expect(sheet.columns.get(2)?.width).toBe(5);
    expect(sheet.frozenRows).toBe(2);
    history.redo();
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('x'));
    expect(sheet.getValue(9, 9)).toEqual(CellValue.empty);
  });
});

describe('CsvOptions.forFileName', () => {
  it('reads a .tsv as tab-separated and anything else as comma-separated', () => {
    expect(CsvOptions.forFileName('data.TSV').delimiter).toBe('\t');
    expect(CsvOptions.forFileName('data.csv').delimiter).toBe(',');
    expect(CsvOptions.forFileName('-').delimiter).toBe(',');
  });
});
