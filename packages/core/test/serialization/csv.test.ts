import { describe, expect, it } from 'vitest';
import { Workbook } from '../../src/model/workbook.ts';
import { CsvOptions, parseCsv, readCsv, writeCsv } from '../../src/serialization/index.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const INVARIANT = { ...CsvOptions.default, locale: 'en-US' };

describe('Csv', () => {
  it('writes rows with LF endings', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.text('a'));
    sheet.setValue(0, 1, CellValue.number(1));
    sheet.setValue(1, 0, CellValue.text('b'));
    sheet.setValue(1, 1, CellValue.number(2));

    expect(writeCsv(workbook, sheet)).toBe('a,1\nb,2\n');
  });

  it('quotes fields with commas, quotes, and newlines', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.text('a,b'));
    sheet.setValue(0, 1, CellValue.text('say "hi"'));
    sheet.setValue(0, 2, CellValue.text('line1\nline2'));

    expect(writeCsv(workbook, sheet)).toBe('"a,b","say ""hi""","line1\nline2"\n');
  });

  it('parses quoted fields and embedded delimiters', () => {
    const rows = parseCsv('a,"b,c","d""e"\n1,2,3\n');

    expect(rows).toHaveLength(2);
    expect(rows[0]).toEqual(['a', 'b,c', 'd"e']);
    expect(rows[1]).toEqual(['1', '2', '3']);
  });

  it('strips BOM and accepts CRLF', () => {
    const rows = parseCsv('﻿a,b\r\nc,d\r\n');

    expect(rows).toHaveLength(2);
    expect(rows[0]).toEqual(['a', 'b']);
  });

  it('infers numbers on import', () => {
    const sheet = readCsv('10,hello,3.5\n', INVARIANT).activeSheet;

    expect(sheet.getValue(0, 0)).toEqual(CellValue.number(10));
    expect(sheet.getValue(0, 1)).toEqual(CellValue.text('hello'));
    expect(sheet.getValue(0, 2)).toEqual(CellValue.number(3.5));
  });

  it('round trips text and numbers', () => {
    const workbook = readCsv('name,score\nAda,99\nBev,87\n', INVARIANT);

    expect(writeCsv(workbook, workbook.activeSheet)).toBe('name,score\nAda,99\nBev,87\n');
  });

  // ---- beyond the C# suite: delimiter options and locale inference ----

  it('honours the semicolon and tab delimiters', () => {
    expect(parseCsv('a;b\n', CsvOptions.semicolon)).toEqual([['a', 'b']]);
    expect(parseCsv('a\tb\n', CsvOptions.tab)).toEqual([['a', 'b']]);
  });

  it('infers numbers with the locale separators', () => {
    const sheet = readCsv('"1.234,5";x\n', { ...CsvOptions.semicolon, locale: 'de-DE' }).activeSheet;
    expect(sheet.getValue(0, 0)).toEqual(CellValue.number(1234.5));
  });

  it('infers dates and gives them a date format', () => {
    const workbook = readCsv('2024-01-15,1/2/2024\n', INVARIANT);
    const sheet = workbook.activeSheet;
    const iso = sheet.find(0, 0)!;
    const local = sheet.find(0, 1)!;

    expect(iso.value).toEqual(CellValue.number(45306));
    expect(workbook.styles.get(iso.styleId).numberFormat.code).toBe('yyyy-mm-dd');
    expect(local.value).toEqual(CellValue.number(45293)); // en-US is month/day/year → 2 Jan 2024
  });

  it('reads day/month/year in a day-first locale', () => {
    const sheet = readCsv('1/2/2024\n', { ...CsvOptions.default, locale: 'en-GB' }).activeSheet;
    expect(sheet.getValue(0, 0)).toEqual(CellValue.number(45323)); // 1 Feb 2024
  });

  it('keeps a lone quote run and an unterminated final record', () => {
    expect(parseCsv('a,b')).toEqual([['a', 'b']]);
    expect(parseCsv('"x"')).toEqual([['x']]);
    expect(parseCsv('')).toEqual([]);
  });
});
