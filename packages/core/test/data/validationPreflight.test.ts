import { describe, expect, it } from 'vitest';
import type { IncomingWrite } from '../../src/data/validationPreflight.ts';
import { ValidationPreflight } from '../../src/data/validationPreflight.ts';
import { dataValidation } from '../../src/data/types.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';

/** Batch validation pre-flight for paste/fill (PRD §8.13). */
const LOCALE = 'en-US';

function sheetWithWholeNumberRule(target: CellRange, min: number, max: number): Worksheet {
  const sheet = new Worksheet('Sheet1');
  sheet.addValidation(
    dataValidation('wholeNumber', target, { operator: 'between', operand1: min, operand2: max }),
  );
  return sheet;
}

const write = (row: number, column: number, value: CellValue, isFormula: boolean): IncomingWrite => ({
  row,
  column,
  value,
  isFormula,
});

describe('ValidationPreflight', () => {
  it('reports cells that violate a stop rule', () => {
    const sheet = sheetWithWholeNumberRule(CellRange.parse('A1:A3'), 1, 10);
    const incoming = [
      write(0, 0, CellValue.number(5), false),
      write(1, 0, CellValue.number(99), false),
      write(2, 0, CellValue.number(3), false),
    ];
    expect(ValidationPreflight.findStopViolations(sheet, incoming, LOCALE)).toEqual([[1, 0]]);
  });

  it('skips formula targets', () => {
    const sheet = sheetWithWholeNumberRule(CellRange.parse('A1:A3'), 1, 10);
    expect(
      ValidationPreflight.findStopViolations(sheet, [write(0, 0, CellValue.number(999), true)], LOCALE),
    ).toEqual([]);
  });

  it('empty values pass', () => {
    const sheet = sheetWithWholeNumberRule(CellRange.parse('A1:A3'), 1, 10);
    expect(
      ValidationPreflight.findStopViolations(sheet, [write(0, 0, CellValue.empty, false)], LOCALE),
    ).toEqual([]);
  });

  it('warning and information rules do not block', () => {
    const sheet = new Worksheet('Sheet1');
    sheet.addValidation(
      dataValidation('wholeNumber', CellRange.parse('A1:A1'), {
        operator: 'between',
        operand1: 1,
        operand2: 10,
        response: 'warning',
      }),
    );
    expect(
      ValidationPreflight.findStopViolations(sheet, [write(0, 0, CellValue.number(99), false)], LOCALE),
    ).toEqual([]);
  });

  it('list rule backed by a range uses the resolver', () => {
    const sheet = new Worksheet('Sheet1');
    sheet.setValue(0, 5, CellValue.text('Red'));
    sheet.setValue(1, 5, CellValue.text('Green'));
    sheet.addValidation(
      dataValidation('list', CellRange.parse('A1:A2'), { listSource: CellRange.parse('F1:F2') }),
    );
    const resolve = (r: CellRange): CellValue[] =>
      [...r.positions()].map(([row, col]) => sheet.getValue(row, col));
    const incoming = [
      write(0, 0, CellValue.text('Green'), false),
      write(1, 0, CellValue.text('Blue'), false),
    ];
    expect(ValidationPreflight.findStopViolations(sheet, incoming, LOCALE, resolve)).toEqual([[1, 0]]);
  });
});
