import { describe, expect, it } from 'vitest';
import { FunctionRegistry } from '../../src/formulas/functions/functionRegistry.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { bool, civil, errorToken, num, text } from './formulaHarness.ts';

const column =
  (col: number, ...values: number[]) =>
  (sheet: Worksheet): void => {
    values.forEach((v, i) => sheet.setValue(i, col, CellValue.number(v)));
  };

const twoColumns = (s: Worksheet): void => {
  column(0, 1, 2, 3, 4, 5)(s);
  column(1, 10, 20, 30, 40, 50)(s);
};

describe('Functions', () => {
  // ---- Math / statistics ------------------------------------------------

  it.each([
    ['SUM(1,2,3)', 6],
    ['PRODUCT(2,3,4)', 24],
    ['AVERAGE(2,4,6)', 4],
    ['MIN(5,2,8)', 2],
    ['MAX(5,2,8)', 8],
    ['ABS(-5)', 5],
    ['INT(3.7)', 3],
    ['INT(-3.2)', -4],
    ['MOD(10,3)', 1],
    ['MOD(-3,2)', 1],
    ['ROUND(3.14159,2)', 3.14],
    ['ROUND(2.5,0)', 3],
    ['ROUND(-2.5,0)', -3],
    ['ROUNDUP(3.1,0)', 4],
    ['ROUNDDOWN(3.9,0)', 3],
    ['POWER(2,10)', 1024],
    ['SQRT(16)', 4],
  ])('math scalar %s = %d', (formula, expected) => {
    expect(num(formula)).toBeCloseTo(expected, 10);
  });

  // A value typed directly as an argument is coerced like arithmetic — booleans become 1/0 and
  // numeric strings parse — unlike values pulled from a range.
  it.each([
    ['SUM(TRUE,FALSE,5)', 6],
    ['AVERAGE(TRUE,FALSE)', 0.5],
    ['MAX(TRUE,0.5)', 1],
    ['COUNT("x",1,TRUE)', 2],
    ['SUM(1,"2")', 3],
  ])('direct arguments coerce booleans and numeric text: %s', (formula, expected) => {
    expect(num(formula)).toBeCloseTo(expected, 10);
  });

  // A direct non-numeric / non-logical argument is an error (a range would ignore it instead).
  it.each(['SUM(1,"x")', 'AND(1,"x")'])('direct non-value argument errors: %s', (formula) => {
    expect(errorToken(formula)).toBe('#VALUE!');
  });

  // Number→text coercion keeps fixed-point across Calc's wide band; integers a double holds
  // exactly print in full, larger ones in scientific notation like Calc.
  it.each([
    ['0.000001&""', '0.000001'],
    ['1000000000000000&""', '1000000000000000'],
    ['1234567890123456&""', '1234567890123456'],
    ['100000000000000000&""', '1E+17'],
    ['2^60&""', '1.15292150460685E+18'],
    ['-(2^60)&""', '-1.15292150460685E+18'],
    ['2^53&""', '9.00719925474099E+15'],
  ])('number to text avoids scientific notation: %s', (formula, expected) => {
    expect(text(formula)).toBe(expected);
  });

  it('SUM over a range ignores text', () => {
    const result = num('SUM(A1:A3)', (s) => {
      s.setValue(0, 0, CellValue.number(1));
      s.setValue(1, 0, CellValue.text('ignored'));
      s.setValue(2, 0, CellValue.number(3));
    });
    expect(result).toBe(4);
  });

  it('COUNT and COUNTA', () => {
    const setup = (s: Worksheet): void => {
      s.setValue(0, 0, CellValue.number(1));
      s.setValue(1, 0, CellValue.text('x'));
      s.setValue(2, 0, CellValue.number(3));
    };
    expect(num('COUNT(A1:A3)', setup)).toBe(2);
    expect(num('COUNTA(A1:A3)', setup)).toBe(3);
  });

  it('AVERAGE of no numbers is #DIV/0!', () => {
    expect(errorToken('AVERAGE(A1:A3)')).toBe('#DIV/0!');
  });

  it('SQRT of a negative is an error', () => {
    expect(errorToken('SQRT(-1)')).toBe('#VALUE!');
  });

  it('SUMIF and COUNTIF', () => {
    expect(num('SUMIF(A1:A5,">2")', column(0, 1, 2, 3, 4, 5))).toBe(12);
    expect(num('COUNTIF(A1:A5,">2")', column(0, 1, 2, 3, 4, 5))).toBe(3);
  });

  it('SUMIF with a separate sum range', () => {
    expect(num('SUMIF(A1:A5,">2",B1:B5)', twoColumns)).toBe(120);
  });

  it('SUMIF reshapes a smaller sum range to the criteria range', () => {
    // B1:B3 is anchored at B1 and reshaped to A1:A5; rows 3,4,5 read B3,B4,B5 = 120.
    expect(num('SUMIF(A1:A5,">2",B1:B3)', twoColumns)).toBe(120);
  });

  it('SUMIF reshapes a single-cell sum range', () => {
    expect(num('SUMIF(A1:A5,">2",B1)', twoColumns)).toBe(120);
  });

  it('SUMIF reshapes an offset sum range', () => {
    // C2:C6 anchored at C2; matched offsets 2,3,4 read C4,C5,C6 = 300+400+500.
    const setup = (s: Worksheet): void => {
      column(0, 1, 2, 3, 4, 5)(s);
      for (let i = 0; i < 5; i++) s.setValue(1 + i, 2, CellValue.number((i + 1) * 100));
    };
    expect(num('SUMIF(A1:A5,">2",C2:C6)', setup)).toBe(1200);
  });

  // ---- Logical ----------------------------------------------------------

  it.each([
    ['IF(1>0,10,20)', 10],
    ['IF(1<0,10,20)', 20],
    ['IFERROR(1/0,42)', 42],
    ['IFERROR(7,42)', 7],
  ])('logical number %s = %d', (formula, expected) => {
    expect(num(formula)).toBeCloseTo(expected, 10);
  });

  it.each([
    ['AND(TRUE,TRUE)', true],
    ['AND(TRUE,FALSE)', false],
    ['OR(FALSE,TRUE)', true],
    ['OR(FALSE,FALSE)', false],
    ['NOT(TRUE)', false],
    ['TRUE()', true],
    ['FALSE()', false],
  ])('logical boolean %s is %s', (formula, expected) => {
    expect(bool(formula)).toBe(expected);
  });

  // ---- Text -------------------------------------------------------------

  it.each([
    ['LEFT("Cellestra",4)', 'Cell'],
    ['RIGHT("Cellestra",3)', 'tra'],
    ['MID("Cellestra",2,4)', 'elle'],
    ['LOWER("ABC")', 'abc'],
    ['UPPER("abc")', 'ABC'],
    ['TRIM("  a  b ")', 'a b'],
    ['SUBSTITUTE("a-b-c","-","+")', 'a+b+c'],
    ['SUBSTITUTE("a-b-c","-","+",2)', 'a-b+c'],
    ['CONCAT("a","b","c")', 'abc'],
    ['CONCATENATE("a",1,TRUE())', 'a1TRUE'],
    ['TEXT(0.25,"0.0%")', '25.0%'],
  ])('text function %s', (formula, expected) => {
    expect(text(formula)).toBe(expected);
  });

  it('LEN and VALUE', () => {
    expect(num('LEN("Cellestra")')).toBe(9);
    expect(num('VALUE("123")')).toBe(123);
  });

  it.each([
    ['VALUE("50%")', 0.5], // a trailing percent divides by 100
    ['VALUE("100%")', 1.0],
    ['VALUE("$100")', 100], // a leading currency symbol is stripped
    ['VALUE("$1.5")', 1.5],
  ])('VALUE parses percent and currency: %s', (formula, expected) => {
    expect(num(formula)).toBeCloseTo(expected, 10);
  });

  // ---- Lookup -----------------------------------------------------------

  it('MATCH exact and approximate', () => {
    const setup = column(0, 10, 20, 30, 40, 50);
    expect(num('MATCH(30,A1:A5,0)', setup)).toBe(3);
    expect(num('MATCH(25,A1:A5,1)', setup)).toBe(2);
  });

  it('MATCH not found is #N/A', () => {
    expect(errorToken('MATCH(99,A1:A5,0)', column(0, 10, 20, 30, 40, 50))).toBe('#N/A');
  });

  it('VLOOKUP exact, INDEX and HLOOKUP', () => {
    const setup = (s: Worksheet): void => {
      s.setValue(0, 0, CellValue.number(1));
      s.setValue(0, 1, CellValue.text('one'));
      s.setValue(1, 0, CellValue.number(2));
      s.setValue(1, 1, CellValue.text('two'));
      s.setValue(2, 0, CellValue.number(3));
      s.setValue(2, 1, CellValue.text('three'));
    };
    expect(text('VLOOKUP(2,A1:B3,2,FALSE)', setup)).toBe('two');
    expect(text('INDEX(A1:B3,3,2)', setup)).toBe('three');
    expect(
      text('HLOOKUP("two",A1:B3,1)', (s) => {
        // For HLOOKUP search the first row.
        s.setValue(0, 0, CellValue.text('one'));
        s.setValue(0, 1, CellValue.text('two'));
        s.setValue(1, 0, CellValue.number(1));
        s.setValue(1, 1, CellValue.number(2));
      }),
    ).toBe('two');
  });

  // ---- Date / time ------------------------------------------------------

  it('DATE round-trips through components', () => {
    expect(num('YEAR(DATE(2026,6,7))')).toBe(2026);
    expect(num('MONTH(DATE(2026,6,7))')).toBe(6);
    expect(num('DAY(DATE(2026,6,7))')).toBe(7);
  });

  it.each([
    ['WEEKDAY(DATE(2000,1,1),1)', 7], // Saturday, Sunday=1 system
    ['WEEKDAY(DATE(2000,1,1),2)', 6], // Monday=1 system
    ['WEEKDAY(DATE(2000,1,1),3)', 5], // Monday=0 system
  ])('WEEKDAY type %s = %d', (formula, expected) => {
    expect(num(formula)).toBe(expected);
  });

  it('NETWORKDAYS counts weekdays', () => {
    // 2026-06-01 (Mon) … 2026-06-05 (Fri) = 5 working days.
    expect(num('NETWORKDAYS(DATE(2026,6,1),DATE(2026,6,5))')).toBe(5);
    // Spanning the weekend to Mon 2026-06-08 adds one more working day.
    expect(num('NETWORKDAYS(DATE(2026,6,1),DATE(2026,6,8))')).toBe(6);
  });

  it('TODAY uses the injected clock', () => {
    const today = num('TODAY()', undefined, civil(2026, 6, 7));
    expect(today).toBe(num('DATE(2026,6,7)'));
  });

  it('the library has the full v1 set', () => {
    const registry = FunctionRegistry.createDefault();
    const expected = [
      ...['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'PRODUCT', 'ABS', 'INT', 'MOD'],
      ...['ROUND', 'ROUNDUP', 'ROUNDDOWN', 'POWER', 'SQRT', 'SUMIF', 'COUNTIF'],
      ...['IF', 'IFERROR', 'AND', 'OR', 'NOT', 'TRUE', 'FALSE'],
      ...['LEFT', 'RIGHT', 'MID', 'LEN', 'LOWER', 'UPPER', 'TRIM', 'SUBSTITUTE', 'TEXT', 'VALUE', 'CONCAT'],
      ...['VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH'],
      ...['TODAY', 'NOW', 'DATE', 'YEAR', 'MONTH', 'DAY', 'WEEKDAY', 'NETWORKDAYS'],
    ];
    for (const name of expected) expect(registry.has(name), `Missing function ${name}`).toBe(true);
  });
});
