import { describe, expect, it } from 'vitest';
import { DateConversion } from '../../src/formulas/evaluation/dateConversion.ts';
import { CellStyles, Fills, Fonts, NumberFormats } from '../../src/styling/cellStyle.ts';
import { Argb } from '../../src/styling/color.ts';
import { NumberFormatCode } from '../../src/styling/numberFormatCode.ts';
import type { NumberFormatKind } from '../../src/styling/numberFormatParser.ts';
import { NumberFormatter } from '../../src/styling/numberFormatter.ts';
import { StylePool } from '../../src/styling/stylePool.ts';
import { CellError } from '../../src/values/cellError.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const INVARIANT = 'en-US';
const fmt = (value: number, code: string): string =>
  NumberFormatter.format(CellValue.number(value), code, INVARIANT);
const bold = () => ({ ...CellStyles.default, font: { ...Fonts.default, bold: true } });

describe('StylePool', () => {
  it('dedupes equal styles', () => {
    const pool = new StylePool();
    const first = pool.intern(bold());
    const second = pool.intern(bold());
    expect(first).toBe(second);
    expect(pool.count).toBe(2); // default + bold
  });

  it('the default is index zero', () => {
    const pool = new StylePool();
    expect(pool.get(StylePool.defaultId)).toEqual(CellStyles.default);
    expect(pool.intern(CellStyles.default)).toBe(0);
  });

  it('mutate derives and interns', () => {
    const pool = new StylePool();
    const id = pool.mutate(StylePool.defaultId, (s) => ({
      ...s,
      fill: Fills.solid(Argb.fromRgb(255, 0, 0)),
    }));
    expect(Fills.hasBackground(pool.get(id).fill)).toBe(true);
    expect(pool.get(id).fill.background).toBe(Argb.fromRgb(255, 0, 0));
  });
});

describe('Argb', () => {
  it.each([
    ['#FF0000', 255, 0, 0],
    ['00FF00', 0, 255, 0],
  ])('parses hex %s', (hex, r, g, b) => {
    expect(Argb.tryFromHex(hex)).toBe(Argb.fromRgb(r, g, b));
  });

  it('round-trips hex', () => {
    expect(Argb.toHex(Argb.fromRgb(0x1a, 0x2b, 0x3c))).toBe('#1A2B3C');
  });
});

describe('NumberFormat', () => {
  it.each<[string, NumberFormatKind]>([
    ['General', 'general'],
    ['0.00', 'number'],
    ['0%', 'percent'],
    ['$#,##0.00', 'currency'],
    ['yyyy-mm-dd', 'date'],
    ['hh:mm:ss', 'time'],
    ['@', 'text'],
  ])('classifies %s as %s', (code, expected) => {
    expect(NumberFormats.kind(NumberFormats.of(code))).toBe(expected);
  });
});

describe('NumberFormatter', () => {
  it.each([
    ['0.00', 3.14159, '3.14'],
    ['0%', 0.25, '25%'],
    ['#,##0', 1234567, '1,234,567'],
    ['General', 42, '42'],
    ['General', 1.5, '1.5'],
  ])('formats numbers: %s of %d', (code, value, expected) => {
    expect(fmt(value, code)).toBe(expected);
  });

  it('formats dates', () => {
    expect(fmt(DateConversion.dateToSerial(2026, 6, 7), 'yyyy-mm-dd')).toBe('2026-06-07');
  });

  // 0.520834 of a day ≈ 12:30:00. AM/PM renders as "PM" (12-hour clock); "mm" between hours and
  // seconds is minutes, not the month.
  it.each([
    ['h:mm:ss AM/PM', 0.520834, '12:30:00 PM'],
    ['h:mm AM/PM', 0.1, '2:24 AM'],
    ['hh:mm:ss', 0.520834, '12:30:00'],
  ])('formats times: %s', (code, value, expected) => {
    expect(fmt(value, code)).toBe(expected);
  });

  // Single-character date/time codes are fields, not standard specifiers. "[h]:mm" shows true
  // elapsed hours.
  it.each([
    ['h', 0.520834, '12'],
    ['s', 0.520834, '0'],
    ['[h]:mm', 0.520834, '12:30'],
    ['[h]:mm', 1.5, '36:00'],
  ])('handles single-char and bracketed time code %s', (code, value, expected) => {
    expect(fmt(value, code)).toBe(expected);
  });

  it('renders date and time components together', () => {
    const serial = DateConversion.dateToSerial(2026, 6, 7) + 0.520834;
    expect(fmt(serial, 'yyyy-mm-dd hh:mm:ss')).toBe('2026-06-07 12:30:00');
  });

  it('passes errors and text through', () => {
    expect(NumberFormatter.format(CellValue.error(CellError.divideByZero()), '0.00', INVARIANT)).toBe(
      '#DIV/0!',
    );
    expect(NumberFormatter.format(CellValue.text('hello'), '0.00', INVARIANT)).toBe('hello');
  });

  // A serial outside the representable date range falls back to a plain number.
  it.each([99999999.0, -800000.0])('falls back on out-of-range date serial %d', (serial) => {
    expect(fmt(serial, 'yyyy-mm-dd')).toBe(NumberFormatter.formatGeneral(serial, INVARIANT));
  });

  // ---- The rest of the format dialect ---------------

  it.each([
    ['0.00;(0.00)', -3.5, '(3.50)'],
    ['0.00;(0.00);"zero"', 0, 'zero'],
    ['[RED]0.00;[BLUE]-0.00', -1, '-1.00'],
    ['#,##0,', 1234567, '1,235'],
    ['0.0,,"M"', 2500000, '2.5M'],
    ['#.##', 3.1, '3.1'],
    ['?.??', 3.1, '3.1 '],
    ['0.00E+00', 12345, '1.23E+04'],
    ['000', 7, '007'],
    ['\\$0', 5, '$5'],
    ['"Total: "0', 5, 'Total: 5'],
    ['[$€-407] #,##0.00', 1234.5, '€ 1,234.50'],
    ['0.0%', 0.123, '12.3%'],
  ])('format %s of %d', (code, value, expected) => {
    expect(fmt(value, code)).toBe(expected);
  });

  it('passes text through unchanged even with a text section', () => {
    expect(NumberFormatter.format(CellValue.text('x'), '0;0;0;"<"@">"', INVARIANT)).toBe('x');
  });

  it('formats month and day names', () => {
    const serial = DateConversion.dateToSerial(2026, 6, 7);
    expect(fmt(serial, 'dddd, mmmm d, yyyy')).toBe('Sunday, June 7, 2026');
    expect(fmt(serial, 'ddd mmm yy')).toBe('Sun Jun 26');
  });

  it('localizes separators and names via Intl', () => {
    expect(NumberFormatter.format(CellValue.number(1234.5), '#,##0.00', 'de-DE')).toBe('1.234,50');
    expect(
      NumberFormatter.format(CellValue.number(DateConversion.dateToSerial(2026, 6, 7)), 'mmmm', 'de-DE'),
    ).toBe('Juni');
    expect(NumberFormatter.formatGeneral(1.5, 'de-DE')).toBe('1,5');
  });

  it('shows tiny numbers in General as scientific, not as 0', () => {
    expect(NumberFormatter.formatGeneral(1e-12, INVARIANT)).toBe('1E-12');
    expect(NumberFormatter.formatGeneral(-2.5e-15, INVARIANT)).toBe('-2.5E-15');
    expect(NumberFormatter.formatGeneral(1.25e-11, 'de-DE')).toBe('1,25E-11');
    expect(NumberFormatter.formatGeneral(1e-10, INVARIANT)).toBe('0.0000000001');
    expect(NumberFormatter.formatGeneral(0.5, INVARIANT)).toBe('0.5');
  });

  it('formats booleans and blanks', () => {
    expect(NumberFormatter.format(CellValue.boolean(true), '0.00', INVARIANT)).toBe('TRUE');
    expect(NumberFormatter.format(CellValue.empty, '0.00', INVARIANT)).toBe('');
  });
});

describe('NumberFormatCode', () => {
  it.each([
    ['0', '0.0'],
    ['0.00', '0.000'],
    ['#,##0.00', '#,##0.000'],
    ['$#,##0.00', '$#,##0.000'],
    ['0.00%', '0.000%'],
    ['0.00E+00', '0.000E+00'], // only the mantissa grows; the exponent's "00" is left alone
    ['General', '#,##0.0'], // General seeds from the grouped number, then steps up
  ])('increase adds a decimal: %s', (code, expected) => {
    expect(NumberFormatCode.increaseDecimals(code)).toBe(expected);
  });

  it.each([
    ['0.000', '0.00'],
    ['0.0', '0'],
    ['0', '0'], // already none — clamps, no negative
    ['#,##0.00', '#,##0.0'],
    ['$#,##0.0', '$#,##0'],
    ['0.00E+00', '0.0E+00'],
    ['General', '#,##0'], // General has no decimals; stays at zero
  ])('decrease removes a decimal: %s', (code, expected) => {
    expect(NumberFormatCode.decreaseDecimals(code)).toBe(expected);
  });

  it.each(['yyyy-mm-dd', 'hh:mm:ss', '@'])('leaves non-numeric code %s untouched', (code) => {
    expect(NumberFormatCode.increaseDecimals(code)).toBe(code);
    expect(NumberFormatCode.decreaseDecimals(code)).toBe(code);
  });
});
