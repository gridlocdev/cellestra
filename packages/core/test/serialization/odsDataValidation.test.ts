import { describe, expect, it } from 'vitest';
import { dataValidation } from '../../src/data/types.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { readOds, writeOds } from '../../src/serialization/index.ts';
import { odsWith, readEntry, roundTrip } from './helpers.ts';

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

  // LibreOffice writes no error message, or one with display="false", when "Show error message
  // when invalid values are entered" is off: any entry is then accepted without a word.
  const ruleWith = (error: string) =>
    odsWith(`<table:content-validations><table:content-validation table:name="val1" table:condition="of:cell-content-is-in-list(&quot;a&quot;;&quot;b&quot;)" table:base-cell-address="Sheet1.A1">${error}</table:content-validation></table:content-validations>
      <table:table table:name="Sheet1"><table:table-row><table:table-cell table:content-validation-name="val1"/></table:table-row></table:table>`);

  it('reads a rule without an error alert as accepting anything quietly', () => {
    for (const error of ['', '<table:error-message table:display="false" table:message-type="stop"/>']) {
      const rule = readOds(ruleWith(error)).findSheet('Sheet1')!.validations[0]!;
      expect(rule.response).toBe('none');
    }
  });

  it('saves a quiet rule without turning its alert on', () => {
    for (const error of ['', '<table:error-message table:display="false" table:message-type="stop"/>']) {
      const saved = writeOds(readOds(ruleWith(error)));
      expect(readEntry(saved, 'content.xml')).not.toMatch(/<table:error-message[^>]*table:display="true"/);
      expect(readOds(saved).findSheet('Sheet1')!.validations[0]!.response).toBe('none');
    }
    const workbook = Workbook.createNew();
    workbook.activeSheet.addValidation(
      dataValidation('list', CellRange.parse('A1'), { listValues: ['a'], response: 'none' }),
    );
    expect(onlyRule(workbook).response).toBe('none');
  });
});
