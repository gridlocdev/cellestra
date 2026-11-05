import { describe, expect, it } from 'vitest';
import { dataValidation } from '../../src/data/types.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { roundTrip } from './helpers.ts';

/**
 * Round-trips ODF `table:content-validation` (PRD §8.13): kind, operator, operands, list values,
 * and the error response survive save → reopen, referenced from validated cells by name.
 */
function onlyRule(workbook: Workbook) {
  const validations = roundTrip(workbook).findSheet('Sheet1')!.validations;
  expect(validations).toHaveLength(1);
  return validations[0]!;
}

describe('OdsDataValidation', () => {
  it('whole number between rule round trips', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.number(5));
    sheet.addValidation(
      dataValidation('wholeNumber', CellRange.parse('A1:A3'), {
        operator: 'between',
        operand1: 1,
        operand2: 10,
        response: 'stop',
        inputMessage: 'Enter 1 to 10',
        errorMessage: 'Out of range',
      }),
    );

    const rule = onlyRule(workbook);
    expect(rule.kind).toBe('wholeNumber');
    expect(rule.operator).toBe('between');
    expect(rule.operand1).toBe(1);
    expect(rule.operand2).toBe(10);
    expect(rule.target.equals(CellRange.parse('A1:A3'))).toBe(true);
    expect(rule.response).toBe('stop');
    expect(rule.inputMessage).toBe('Enter 1 to 10');
    expect(rule.errorMessage).toBe('Out of range');
  });

  it('decimal relational rule round trips', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.addValidation(
      dataValidation('decimal', CellRange.parse('B2'), {
        operator: 'greaterOrEqual',
        operand1: 2.5,
        response: 'warning',
      }),
    );

    const rule = onlyRule(workbook);
    expect(rule.kind).toBe('decimal');
    expect(rule.operator).toBe('greaterOrEqual');
    expect(rule.operand1).toBe(2.5);
    expect(rule.target.equals(CellRange.parse('B2'))).toBe(true);
    expect(rule.response).toBe('warning');
  });

  it('text length between rule round trips', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.addValidation(
      dataValidation('textLength', CellRange.parse('A1'), { operator: 'between', operand1: 2, operand2: 8 }),
    );

    const rule = onlyRule(workbook);
    expect(rule.kind).toBe('textLength');
    expect(rule.operator).toBe('between');
    expect(rule.operand1).toBe(2);
    expect(rule.operand2).toBe(8);
  });

  it('list rule with literal values round trips', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.addValidation(
      dataValidation('list', CellRange.parse('A1:A2'), {
        listValues: ['Red', 'Green', 'Blue'],
        response: 'information',
      }),
    );

    const rule = onlyRule(workbook);
    expect(rule.kind).toBe('list');
    expect(rule.listValues).toEqual(['Red', 'Green', 'Blue']);
    expect(rule.response).toBe('information');
  });

  it('list rule with quoted value round trips', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.addValidation(
      dataValidation('list', CellRange.parse('A1'), { listValues: ['a;b', 'say "hi"'] }),
    );

    expect(onlyRule(workbook).listValues).toEqual(['a;b', 'say "hi"']);
  });

  it('list rule with a cell range source round trips', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.addValidation(
      dataValidation('list', CellRange.parse('A1:A3'), {
        listSource: CellRange.parse('E1:E8'),
        response: 'stop',
      }),
    );

    const rule = onlyRule(workbook);
    expect(rule.kind).toBe('list');
    expect(rule.listValues ?? null).toBeNull();
    expect(rule.listSource).toBeTruthy();
    expect(rule.listSource!.left).toBe(4); // column E → index 4
    expect(rule.listSource!.top).toBe(0); // row 1 → index 0
    expect(rule.listSource!.bottom).toBe(7); // row 8 → index 7
  });

  it('no validation emits no block and reopens clean', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.setValue(0, 0, CellValue.number(1));

    expect(roundTrip(workbook).findSheet('Sheet1')!.validations).toHaveLength(0);
  });

  it('validation on blank cells still round trips', () => {
    // No cell values at all — the validated cells are blank and past any used range.
    const workbook = Workbook.createNew();
    workbook.activeSheet.addValidation(
      dataValidation('list', CellRange.parse('C5:C7'), { listValues: ['Yes', 'No'] }),
    );

    const rule = onlyRule(workbook);
    expect(rule.target.equals(CellRange.parse('C5:C7'))).toBe(true);
    expect(rule.listValues).toEqual(['Yes', 'No']);
  });
});
