import { describe, expect, it } from 'vitest';
import { sheetDeletionImpact } from '../../src/editing/deleteSheetCommand.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

function twoSheets() {
  const workbook = Workbook.createNew();
  const first = workbook.activeSheet;
  const second = workbook.addSheet('Data');
  return { workbook, first, second };
}

describe('sheetDeletionImpact', () => {
  it('calls a sheet with only formatting empty, and one with a value not', () => {
    const { workbook, second } = twoSheets();
    expect(sheetDeletionImpact(workbook, 'Data')).toEqual({ empty: true, formulas: 0, otherReferences: 0 });
    second.setValue(4, 2, CellValue.number(1));
    expect(sheetDeletionImpact(workbook, 'Data').empty).toBe(false);
  });

  it('counts only formulas elsewhere that actually read the sheet', () => {
    const { workbook, first, second } = twoSheets();
    second.setValue(0, 0, CellValue.number(5));
    first.setFormula(0, 0, 'Data.A1*2');
    first.setFormula(1, 0, 'SUM($Data.A1:B3)');
    first.setFormula(2, 0, 'A1+1');
    // A formula on the sheet itself goes with it; it breaks nothing.
    second.setFormula(1, 0, 'A1');
    expect(sheetDeletionImpact(workbook, 'Data')).toEqual({ empty: false, formulas: 2, otherReferences: 0 });
    expect(sheetDeletionImpact(workbook, first.name).formulas).toBe(0);
  });

  it('counts a formula that reads the sheet through a named range', () => {
    const { workbook, first } = twoSheets();
    workbook.addNamedRange({ name: 'Rates', target: CellRange.parse('Data.A1:A3') });
    first.setFormula(0, 0, 'SUM(Rates)');
    expect(sheetDeletionImpact(workbook, 'Data').formulas).toBe(1);
  });
});
