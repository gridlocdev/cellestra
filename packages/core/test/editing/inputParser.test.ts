import { describe, expect, it } from 'vitest';
import { InputParser } from '../../src/editing/inputParser.ts';
import { DateConversion } from '../../src/formulas/evaluation/dateConversion.ts';
import { NumberFormats } from '../../src/styling/cellStyle.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const EN_US = 'en-US';
const serial = (year: number, month: number, day: number): number =>
  DateConversion.toSerial({ year, month, day, hour: 0, minute: 0, second: 0, millisecond: 0 });

describe('InputParser', () => {
  it.each([
    ["'007", '007'],
    ["'=notaformula", '=notaformula'],
    ["'5", '5'],
    ["'TRUE", 'TRUE'],
    ["''double", "'double"], // only the first apostrophe is the flag
  ])('leading apostrophe forces text: %s', (typed, stored) => {
    const parsed = InputParser.parse(typed, EN_US);
    expect(parsed.formula).toBeNull();
    expect(parsed.value).toEqual(CellValue.text(stored)); // apostrophe stripped, not displayed
  });

  it('bare number still parses as number', () => {
    const v = InputParser.parse('007', EN_US).value;
    expect(v.kind).toBe('number');
    expect(CellValue.asNumber(v)).toBe(7);
  });

  it('slash date coerces to serial with date format', () => {
    const parsed = InputParser.parse('6/7/2026', EN_US);
    expect(parsed.value.kind).toBe('number');
    expect(CellValue.asNumber(parsed.value)).toBeCloseTo(serial(2026, 6, 7), 9);
    expect(parsed.suggestedFormat).not.toBeNull();
    expect(NumberFormats.kind(parsed.suggestedFormat!)).toBe('date');
  });

  it('ISO date coerces to serial', () => {
    const parsed = InputParser.parse('2026-06-07', EN_US);
    expect(parsed.value.kind).toBe('number');
    expect(CellValue.asNumber(parsed.value)).toBeCloseTo(serial(2026, 6, 7), 9);
    expect(NumberFormats.kind(parsed.suggestedFormat!)).toBe('date');
  });

  it('date parsing follows locale day-month order', () => {
    // 03/04/2026 is 3 April in de-DE, 4 March in en-US (#38).
    expect(CellValue.asNumber(InputParser.parse('03/04/2026', 'de-DE').value)).toBeCloseTo(
      serial(2026, 4, 3),
      9,
    );
    expect(CellValue.asNumber(InputParser.parse('03/04/2026', EN_US).value)).toBeCloseTo(
      serial(2026, 3, 4),
      9,
    );
  });

  // Two-digit years resolve on LibreOffice Calc's 1930–2029 window: 00–29 → 2000–2029, 30–99 → 1930–1999.
  it.each([
    ['6/7/26', 2026],
    ['6/7/29', 2029],
    ['6/7/30', 1930],
    ['6/7/99', 1999],
    ['6/7/00', 2000],
  ])('two-digit year uses the Calc 1930 window: %s', (typed, expectedYear) => {
    const parsed = InputParser.parse(typed, EN_US);
    expect(parsed.value.kind).toBe('number');
    expect(CellValue.asNumber(parsed.value)).toBeCloseTo(serial(expectedYear, 6, 7), 9);
  });

  it('plain text and codes are not coerced to dates', () => {
    expect(InputParser.parse('May', EN_US).value.kind).toBe('text'); // month name alone stays text
    expect(InputParser.parse('part-2', EN_US).value.kind).toBe('text'); // hyphenated token stays text
    expect(InputParser.parse('hello', EN_US).suggestedFormat).toBeNull();
  });
});
