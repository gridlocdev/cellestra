import { describe, expect, it } from 'vitest';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellError } from '../../src/values/cellError.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { approxEqual } from '../../src/formulas/evaluation/valueComparer.ts';
import { bool, errorToken, num, text } from './formulaHarness.ts';

const column =
  (values: readonly CellValue[], col = 0) =>
  (sheet: Worksheet) =>
    values.forEach((v, row) => sheet.setValue(row, col, v));
const n = CellValue.number;
const t = CellValue.text;
const err = CellValue.error(CellError.divideByZero());

describe('exact MATCH/VLOOKUP use value equality, not criteria', () => {
  it('a text lookup that looks like a criteria finds only that text', () => {
    const setup = column([n(5), t('>3'), n(1)]);
    expect(num('MATCH(">3", A1:A3, 0)', setup)).toBe(2);
    expect(errorToken('MATCH("=5", A1:A3, 0)', setup)).toBe('#N/A');
  });

  it('text is never coerced to a number, nor a number to text', () => {
    expect(num('MATCH("5", A1:A3, 0)', column([n(5), t('5'), n(1)]))).toBe(2);
    expect(num('MATCH(5, A1:A3, 0)', column([t('5'), n(5), n(1)]))).toBe(2);
    expect(errorToken('MATCH("5", A1:A2, 0)', column([n(5), n(1)]))).toBe('#N/A');
  });

  it('text matches case-insensitively and with wildcards', () => {
    const setup = column([t('apple'), t('Banana'), t('cherry')]);
    expect(num('MATCH("BANANA", A1:A3, 0)', setup)).toBe(2);
    expect(num('MATCH("ch*", A1:A3, 0)', setup)).toBe(3);
    expect(text('VLOOKUP("b?nana", A1:A3, 1, FALSE())', setup)).toBe('Banana');
  });

  it('an error cell in the lookup vector does not match and does not stop the search', () => {
    const setup = column([err, n(7)]);
    expect(num('MATCH(7, A1:A2, 0)', setup)).toBe(2);
    expect(num('VLOOKUP(7, A1:A2, 1, FALSE())', setup)).toBe(7);
  });
});

describe('approximate MATCH/VLOOKUP skip cells of another type', () => {
  it('a text header above sorted numbers is skipped, not the boundary', () => {
    const setup = column([t('Amount'), n(10), n(20), n(30)]);
    expect(num('MATCH(25, A1:A4, 1)', setup)).toBe(3);
    expect(num('VLOOKUP(25, A1:A4, 1)', setup)).toBe(20);
  });

  it('errors in the vector are skipped too', () => {
    expect(num('MATCH(25, A1:A3, 1)', column([n(10), err, n(20)]))).toBe(3);
  });
});

describe('COUNTIF/SUMIF ignore error cells in the criteria range', () => {
  it('an error never matches', () => {
    const setup = (sheet: Worksheet) => {
      column([n(1), err, n(3)])(sheet);
      column([n(10), n(20), n(30)], 1)(sheet);
    };
    expect(num('COUNTIF(A1:A3, ">0")', setup)).toBe(2);
    expect(num('SUMIF(A1:A3, ">0", B1:B3)', setup)).toBe(40);
    expect(num('COUNTIF(A1:A3, "<>1")', setup)).toBe(1);
  });
});

describe('rounding', () => {
  it('pre-rounds the scaled value to 15 significant digits', () => {
    expect(num('ROUND(1.005, 2)')).toBe(1.01);
    expect(num('ROUND(-1.005, 2)')).toBe(-1.01);
    expect(num('ROUNDDOWN(4.35, 2)')).toBe(4.35);
    expect(num('ROUNDUP(4.35, 2)')).toBe(4.35);
    expect(num('ROUND(2.675, 2)')).toBe(2.68);
  });

  it('clamps out-of-range digits; a non-finite result is an error', () => {
    expect(num('ROUND(1E+10, 400)')).toBe(1e10);
    expect(num('ROUND(1E+300, 300)')).toBe(1e300);
    expect(num('ROUND(123.456, -400)')).toBe(0);
    expect(num('ROUNDDOWN(123.456, -400)')).toBe(0);
    expect(num('ROUNDUP(1.5, -308)')).toBe(1e308);
    // Cellestra has no #NUM!; non-finite results are #VALUE! throughout.
    expect(errorToken('ROUNDUP(1.5, -400)')).toBe('#VALUE!');
  });
});

describe('numeric equality is approximate (relative 2^-48)', () => {
  it('0.1+0.2 = 0.3', () => {
    expect(bool('0.1+0.2=0.3')).toBe(true);
    expect(bool('0.1+0.2<>0.3')).toBe(false);
    expect(bool('0.1+0.2>0.3')).toBe(false);
    expect(bool('1=1.0000001')).toBe(false);
  });

  it('criteria and lookups use it too', () => {
    const setup = (sheet: Worksheet) => sheet.setFormula(0, 0, '0.1+0.2', n(0.1 + 0.2));
    expect(num('COUNTIF(A1:A1, 0.3)', setup)).toBe(1);
    expect(num('COUNTIF(A1:A1, "=0.3")', setup)).toBe(1);
    expect(num('MATCH(0.3, A1:A1, 0)', setup)).toBe(1);
  });

  it('approxEqual keeps zero exact', () => {
    expect(approxEqual(0, 1e-300)).toBe(false);
    expect(approxEqual(1, 1 + 2 ** -50)).toBe(true);
  });
});

describe('booleans compare as numbers', () => {
  it('1=TRUE and 0=FALSE (LibreOffice)', () => {
    expect(bool('1=TRUE()')).toBe(true);
    expect(bool('0=FALSE()')).toBe(true);
    expect(bool('TRUE()>0')).toBe(true);
    expect(bool('2=TRUE()')).toBe(false);
  });
});
