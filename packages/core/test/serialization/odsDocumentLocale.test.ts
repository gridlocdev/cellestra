import { describe, expect, it } from 'vitest';
import { CellRange } from '../../src/model/cellRange.ts';
import { getHostLocale } from '../../src/model/documentSettings.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { DataEditCommands } from '../../src/editing/dataEditCommands.ts';
import { fillFromCsv } from '../../src/serialization/csv/csvReader.ts';
import { CsvOptions, readOds } from '../../src/serialization/index.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { odsWith, roundTrip } from './helpers.ts';

const TABLE =
  '<table:table table:name="Sheet1"><table:table-row><table:table-cell/></table:table-row></table:table>';

const STYLE_NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"';

/** A styles.xml whose default styles are `defaults`, as LibreOffice writes them. */
const stylesXml = (defaults: string): string =>
  `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${STYLE_NS} office:version="1.3"><office:styles>${defaults}</office:styles></office:document-styles>`;

const defaultStyle = (family: string, language: string, country?: string): string =>
  `<style:default-style style:family="${family}"><style:text-properties fo:language="${language}"${
    country === undefined ? '' : ` fo:country="${country}"`
  } style:language-asian="zh" style:country-asian="CN"/></style:default-style>`;

const SETTINGS_NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"';

const settingsXml = (locale: string): string =>
  `<?xml version="1.0" encoding="UTF-8"?><office:document-settings ${SETTINGS_NS}><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item config:name="CellestraLocale" config:type="string">${locale}</config:config-item></config:config-item-set></office:settings></office:document-settings>`;

function open(parts: Record<string, string>): Workbook {
  return readOds(odsWith(TABLE, { extra: parts }));
}

describe('document locale from the .ods default style', () => {
  it('reads fo:language and fo:country of the default cell style', () => {
    const wb = open({ 'styles.xml': stylesXml(defaultStyle('table-cell', 'de', 'DE')) });
    expect(wb.settings.locale).toBe('de-DE');
    expect(wb.settings.effectiveLocale).toBe('de-DE');
  });

  it('takes the language alone when there is no country', () => {
    expect(open({ 'styles.xml': stylesXml(defaultStyle('table-cell', 'fr')) }).settings.locale).toBe('fr');
  });

  it('prefers the cell default style over the paragraph one', () => {
    const styles = stylesXml(defaultStyle('paragraph', 'sv', 'SE') + defaultStyle('table-cell', 'nl', 'NL'));
    expect(open({ 'styles.xml': styles }).settings.locale).toBe('nl-NL');
  });

  it('falls back to another default style that names a language', () => {
    const styles = stylesXml(defaultStyle('paragraph', 'sv', 'SE'));
    expect(open({ 'styles.xml': styles }).settings.locale).toBe('sv-SE');
  });

  it.each([
    ['zxx', 'none'],
    ['none', undefined],
    ['not a language', 'XX'],
  ])('ignores language %s', (language, country) => {
    const wb = open({ 'styles.xml': stylesXml(defaultStyle('table-cell', language, country)) });
    expect(wb.settings.locale).toBeNull();
    expect(wb.settings.effectiveLocale).toBe(getHostLocale());
  });

  it('keeps the host locale when the file names none', () => {
    expect(open({}).settings.locale).toBeNull();
  });

  it('a locale chosen in Cellestra (settings.xml) wins over the default style', () => {
    const wb = open({
      'styles.xml': stylesXml(defaultStyle('table-cell', 'de', 'DE')),
      'settings.xml': settingsXml('en-GB'),
    });
    expect(wb.settings.locale).toBe('en-GB');
  });

  it('a blank Cellestra locale keeps the default style one', () => {
    const wb = open({
      'styles.xml': stylesXml(defaultStyle('table-cell', 'de', 'DE')),
      'settings.xml': settingsXml(''),
    });
    expect(wb.settings.locale).toBe('de-DE');
  });

  it('round-trips', () => {
    const wb = open({ 'styles.xml': stylesXml(defaultStyle('table-cell', 'de', 'DE')) });
    expect(roundTrip(wb).settings.locale).toBe('de-DE');
  });
});

describe('the document locale where a locale was assumed', () => {
  it('CSV import reads numbers in it', () => {
    const wb = open({ 'styles.xml': stylesXml(defaultStyle('table-cell', 'de', 'DE')) });
    fillFromCsv(wb, wb.activeSheet, '"1,5";x\n', CsvOptions.semicolon);
    expect(CellValue.asNumber(wb.activeSheet.getValue(0, 0))).toBe(1.5);
  });

  it('sorting collates text in it', () => {
    const wb = Workbook.createNew();
    wb.settings.locale = 'sv-SE'; // Swedish sorts ö after z
    const sheet = wb.activeSheet;
    sheet.setValue(0, 0, CellValue.text('ö'));
    sheet.setValue(1, 0, CellValue.text('z'));
    DataEditCommands.sortByColumn(sheet.name, CellRange.fromBounds(0, 0, 1, 0), 0, true, false).apply(wb);
    expect([sheet.getValue(0, 0), sheet.getValue(1, 0)].map(CellValue.toInvariantString)).toEqual(['z', 'ö']);
  });

  it('text to columns reads numbers in it', () => {
    const wb = Workbook.createNew();
    wb.settings.locale = 'de-DE';
    const sheet = wb.activeSheet;
    sheet.setValue(0, 0, CellValue.text('1,5;x'));
    DataEditCommands.textToColumns(sheet, CellRange.fromBounds(0, 0, 0, 0), ';').apply(wb);
    expect(CellValue.asNumber(sheet.getValue(0, 0))).toBe(1.5);
  });
});
