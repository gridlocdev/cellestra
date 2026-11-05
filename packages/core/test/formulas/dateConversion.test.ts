import { describe, expect, it } from 'vitest';
import { DateConversion } from '../../src/formulas/evaluation/dateConversion.ts';
import { civil } from './formulaHarness.ts';

/**
 * Locks in the ODF/Calc date-serial conventions: epoch 1899-12-30 and — unlike the Excel 1900
 * system — no phantom 29 Feb 1900 leap day, so naive serial math stays aligned with LibreOffice.
 */
describe('DateConversion', () => {
  it('epoch is serial zero', () => {
    expect(DateConversion.toSerial(civil(1899, 12, 30))).toBeCloseTo(0, 9);
    expect(DateConversion.toSerial(civil(1899, 12, 31))).toBeCloseTo(1, 9);
  });

  it('has no phantom 1900 leap day', () => {
    const feb28 = DateConversion.dateToSerial(1900, 2, 28);
    const mar1 = DateConversion.dateToSerial(1900, 3, 1);
    expect(mar1 - feb28).toBeCloseTo(1, 9);
    expect(DateConversion.dateToSerial(1900, 2, 29)).toBeCloseTo(mar1, 9); // overflow rolls forward
  });

  it('serial round-trips with a time fraction', () => {
    const noon = civil(2026, 6, 7, 12);
    const serial = DateConversion.toSerial(noon);
    expect(serial - Math.floor(serial)).toBeCloseTo(0.5, 9); // noon = half a day
    expect(DateConversion.fromSerial(serial)).toEqual(noon);
  });

  it('date serial matches LibreOffice for a known modern date', () => {
    // LibreOffice Calc: DATE(2024,1,1) → serial 45292 (null date 1899-12-30).
    expect(DateConversion.dateToSerial(2024, 1, 1)).toBeCloseTo(45292, 9);
  });

  it('DATE treats three- or four-digit years literally like LibreOffice', () => {
    const serial = DateConversion.dateToSerial(108, 1, 2);
    expect(DateConversion.fromSerial(serial).year).toBe(108);
  });

  it.each([
    [30, 1930], // 30–99 → 1930–1999
    [99, 1999],
    [0, 2000], // 00–29 → 2000–2029
    [29, 2029],
  ])('DATE resolves two-digit year %d on the 1930–2029 window', (year, expectedYear) => {
    const serial = DateConversion.dateToSerial(year, 1, 1);
    expect(DateConversion.fromSerial(serial).year).toBe(expectedYear);
  });

  it.each([
    // In range: a representable serial converts.
    [0.0, true],
    [45292.0, true],
    // Out of range / non-finite: no throw, just a failed conversion so formatting can fall back.
    [99999999.0, false],
    [-800000.0, false],
    [NaN, false],
    [Infinity, false],
  ])('tryToDateTime(%d) never throws and converts: %s', (serial, expected) => {
    expect(DateConversion.tryToDateTime(serial, false) !== null).toBe(expected);
  });
});
