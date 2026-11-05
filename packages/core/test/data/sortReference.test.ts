import { describe, expect, it } from 'vitest';
import { sortSpec, SortEngine } from '../../src/data/sortEngine.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { numberAt } from './helpers.ts';

describe('SortReference', () => {
  it('per-row formula references follow the sorted row', () => {
    const sheet = new Worksheet('S');
    sheet.setValue(0, 0, CellValue.text('Key'));
    sheet.setValue(0, 1, CellValue.text('Value'));
    sheet.setValue(0, 2, CellValue.text('Doubled'));
    sheet.setValue(1, 0, CellValue.number(3));
    sheet.setValue(1, 1, CellValue.number(30));
    sheet.setFormula(1, 2, 'B2*2', CellValue.empty);
    sheet.setValue(2, 0, CellValue.number(1));
    sheet.setValue(2, 1, CellValue.number(10));
    sheet.setFormula(2, 2, 'B3*2', CellValue.empty);
    sheet.setValue(3, 0, CellValue.number(2));
    sheet.setValue(3, 1, CellValue.number(20));
    sheet.setFormula(3, 2, 'B4*2', CellValue.empty);

    SortEngine.sort(sheet, CellRange.parse('A1:C4'), sortSpec([{ column: 0, ascending: true }]));

    expect(numberAt(sheet, 1, 0)).toBe(1);
    expect(sheet.find(1, 2)?.formula).toBe('B2*2');
    expect(numberAt(sheet, 2, 0)).toBe(2);
    expect(sheet.find(2, 2)?.formula).toBe('B3*2');
    expect(numberAt(sheet, 3, 0)).toBe(3);
    expect(sheet.find(3, 2)?.formula).toBe('B4*2');
  });

  it('absolute and out-of-range references are unchanged', () => {
    const sheet = new Worksheet('S');
    sheet.setValue(0, 0, CellValue.number(2));
    sheet.setValue(1, 0, CellValue.number(1));
    sheet.setFormula(0, 1, 'A1+$A$10', CellValue.empty);
    sheet.setFormula(1, 1, 'A2+$A$10', CellValue.empty);

    SortEngine.sort(sheet, CellRange.parse('A1:B2'), sortSpec([{ column: 0, ascending: true }], false));

    expect(numberAt(sheet, 0, 0)).toBe(1);
    expect(sheet.find(0, 1)?.formula).toBe('A1+$A$10');
  });
});
