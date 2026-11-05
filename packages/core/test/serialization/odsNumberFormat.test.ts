import { describe, expect, it } from 'vitest';
import { Workbook } from '../../src/model/workbook.ts';
import { readOds, writeOds } from '../../src/serialization/index.ts';
import { CellStyles, NumberFormats, type NumberFormat } from '../../src/styling/cellStyle.ts';
import type { NumberFormatKind } from '../../src/styling/numberFormatParser.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { readEntry } from './helpers.ts';

/**
 * Number formats through open → save → reopen. ODF carries a currency symbol only in
 * `number:currency-style`'s `number:currency-symbol`.
 */
function withFormat(code: string): Workbook {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  sheet.setValue(0, 0, CellValue.number(1234.5));
  sheet.getOrCreate(0, 0).styleId = workbook.styles.intern({
    ...CellStyles.default,
    numberFormat: NumberFormats.of(code),
  });
  return workbook;
}

function roundTripFormat(code: string): NumberFormat {
  const reopened = readOds(writeOds(withFormat(code)));
  return reopened.styles.get(reopened.findSheet('Sheet1')!.find(0, 0)!.styleId).numberFormat;
}

function contentXml(code: string): string {
  return readEntry(writeOds(withFormat(code)), 'content.xml');
}

describe('OdsNumberFormat', () => {
  it('currency survives as currency', () => {
    const format = roundTripFormat('$#,##0.00');

    expect(format.code).toBe('$#,##0.00');
    expect(NumberFormats.kind(format)).toBe('currency');
  });

  it('currency is written as an ODF currency style', () => {
    const xml = contentXml('$#,##0.00');

    expect(xml).toContain('currency-style');
    expect(xml).toContain('currency-symbol');
  });

  it('a trailing currency symbol stays trailing', () => {
    const format = roundTripFormat('#,##0.00 €');

    expect(format.code).toBe('#,##0.00 €');
    expect(NumberFormats.kind(format)).toBe('currency');
  });

  it('currency without grouping survives', () => {
    const format = roundTripFormat('£0.00');

    expect(format.code).toBe('£0.00');
    expect(NumberFormats.kind(format)).toBe('currency');
  });

  it.each<[string, NumberFormatKind]>([
    ['0.00', 'number'],
    ['#,##0.00', 'number'],
    ['0.00%', 'percent'],
    ['yyyy-mm-dd', 'date'],
  ])('the other formats still survive: %s', (code, kind) => {
    const format = roundTripFormat(code);

    expect(format.code).toBe(code);
    expect(NumberFormats.kind(format)).toBe(kind);
  });
});
