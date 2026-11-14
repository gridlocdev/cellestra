import { describe, expect, it } from 'vitest';
import { DuplicateRemover } from '../../src/data/dataTools.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { textAt } from './helpers.ts';

describe('Remove Duplicates ignores case, like Calc’s and Excel’s defaults', () => {
  it('drops a row that differs from an earlier one only in case', () => {
    const sheet = new Worksheet('Sheet1');
    ['Name', 'Apple', 'APPLE', 'apple', 'Pear'].forEach((t, r) => sheet.setValue(r, 0, CellValue.text(t)));
    const removed = DuplicateRemover.remove(sheet, CellRange.parse('A1:A5'), [0], true);
    expect(removed).toBe(2);
    expect(textAt(sheet, 1, 0)).toBe('Apple');
    expect(textAt(sheet, 2, 0)).toBe('Pear');
    expect(sheet.getValue(3, 0).kind).toBe('empty');
  });
});
