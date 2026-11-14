import { ChartDocument, createChart, createSeries } from '@cellestra/charts';
import { describe, expect, it } from 'vitest';
import { dataValidation } from '../../src/data/types.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { MergedRegion } from '../../src/model/mergedRegion.ts';
import { Workbook } from '../../src/model/workbook.ts';
import {
  CsvOptions,
  OdsPackage,
  OdsXmlNames,
  readCsv,
  readOds,
  readOdsWithPackage,
  writeOds,
} from '../../src/serialization/index.ts';
import { CellStyles, NumberFormats } from '../../src/styling/cellStyle.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { readEntry, roundTrip } from './helpers.ts';

/**
 * Interoperability with what LibreOffice writes and reads: each case is a file shape LibreOffice
 * produces, or a reading of Cellestra's output LibreOffice makes, that once lost or changed data.
 */

const NS = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"',
  'xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"',
  'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"',
].join(' ');

interface Parts {
  readonly automaticStyles?: string;
  readonly styles?: string;
  readonly settings?: string;
  readonly extra?: Record<string, string>;
}

function ods(body: string, parts: Parts = {}): Uint8Array {
  const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3"><office:automatic-styles>${parts.automaticStyles ?? ''}</office:automatic-styles><office:body><office:spreadsheet>${body}</office:spreadsheet></office:body></office:document-content>`;
  const pkg = new OdsPackage();
  pkg.setText(OdsXmlNames.mimetypePart, OdsXmlNames.mimeType);
  pkg.setText(OdsXmlNames.contentPart, content);
  if (parts.styles !== undefined) pkg.setText(OdsXmlNames.stylesPart, parts.styles);
  if (parts.settings !== undefined) pkg.setText(OdsXmlNames.settingsPart, parts.settings);
  for (const [name, text] of Object.entries(parts.extra ?? {})) pkg.setText(name, text);
  return pkg.write();
}

function resave(bytes: Uint8Array): Uint8Array {
  const { workbook, package: pkg } = readOdsWithPackage(bytes);
  return writeOds(workbook, pkg);
}

const content = (bytes: Uint8Array): string => readEntry(bytes, OdsXmlNames.contentPart);

describe('chart sheets', () => {
  const chartXml = (): string => {
    const chart = createChart({ type: 'bar' });
    chart.series.push(createSeries({ valuesReference: 'Data.A1:A3' }));
    return ChartDocument.buildContent(chart);
  };
  const frame = (n: number, x: string): string =>
    `<draw:frame draw:name="c${n}" svg:width="8cm" svg:height="6cm" svg:x="${x}" svg:y="1cm"><draw:object xlink:href="./Object ${n}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
  const data = `<table:table table:name="Data"><table:table-column/><table:table-row><table:table-cell office:value-type="float" office:value="1"/></table:table-row></table:table>`;
  const objects = { 'Object 1/content.xml': chartXml(), 'Object 2/content.xml': chartXml() };

  it('a sheet of shapes holding charts is not a chart sheet, and keeps every shape', () => {
    const dash = `<table:table table:name="Dash"><table:shapes>${frame(1, '1cm')}${frame(2, '10cm')}<draw:rect svg:width="1cm" svg:height="1cm" svg:x="0cm" svg:y="9cm"/></table:shapes><table:table-column/><table:table-row><table:table-cell/></table:table-row></table:table>`;
    const bytes = ods(data + dash, { extra: objects });
    expect(readOds(bytes).findSheet('Dash')!.isChartSheet).toBe(false);

    const saved = content(resave(bytes));
    const dashXml = saved.slice(saved.indexOf('table:name="Dash"'));
    expect(dashXml).toContain('./Object 1');
    expect(dashXml).toContain('./Object 2');
    expect(dashXml).toContain('draw:rect');
  });

  it('a chart in A1 beside other cell content is preserved, not modeled', () => {
    const sheet = `<table:table table:name="Dash"><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell>${frame(1, '0cm')}</table:table-cell><table:table-cell><draw:rect svg:width="1cm" svg:height="1cm" svg:x="0cm" svg:y="0cm"/></table:table-cell></table:table-row></table:table>`;
    const bytes = ods(data + sheet, { extra: objects });
    expect(readOds(bytes).findSheet('Dash')!.isChartSheet).toBe(false);
    expect(content(resave(bytes))).toContain('draw:rect');
  });

  it('the frame Cellestra writes alone in A1 is still read as a chart sheet', () => {
    // Cellestra marks the frames of the chart sheets it writes; an unmarked chart is another editor's.
    const marked = frame(1, '0cm').replace(
      '<draw:frame ',
      '<draw:frame xmlns:cellestra="urn:cellestra:names:table:1.0" cellestra:chart-sheet="true" ',
    );
    const sheet = `<table:table table:name="Chart"><table:table-column/><table:table-row><table:table-cell>${marked}</table:table-cell></table:table-row></table:table>`;
    expect(readOds(ods(data + sheet, { extra: objects })).findSheet('Chart')!.isChartSheet).toBe(true);
  });
});

describe('text is written without indentation', () => {
  it('a comment whose paragraph holds only spans reads back as LibreOffice will read it', () => {
    const body = `<table:table table:name="S"><table:table-column/><table:table-row><table:table-cell office:value-type="string"><office:annotation><dc:creator>me</dc:creator><text:p><text:span>Hello</text:span><text:span>World</text:span></text:p></office:annotation><text:p>x</text:p></table:table-cell></table:table-row></table:table>`;
    const saved = content(resave(ods(body)));
    expect(saved).toContain('<text:p><text:span>Hello</text:span><text:span>World</text:span></text:p>');
    expect(saved).not.toMatch(/>\n\s*</);
  });

  it('a Windows line end in a cell is a line break', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.setValue(0, 0, CellValue.text('a\r\nb'));
    const saved = content(writeOds(workbook));
    expect(saved).toContain('<text:p>a</text:p><text:p>b</text:p>');
    expect(roundTrip(workbook).activeSheet.getValue(0, 0)).toEqual(CellValue.text('a\nb'));
  });
});

describe('LibreOffice row padding', () => {
  const optimal = `<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.178in" fo:break-before="auto" style:use-optimal-row-height="true"/></style:style>`;
  const custom = `<style:style style:name="ro2" style:family="table-row"><style:table-row-properties style:row-height="0.5in" style:use-optimal-row-height="false"/></style:style>`;

  it('optimal-height padding (data, a run, then one last row) creates no row models', () => {
    const body = `<table:table table:name="S"><table:table-column table:number-columns-repeated="1024"/><table:table-row table:style-name="ro1"><table:table-cell office:value-type="float" office:value="1"/><table:table-cell table:number-columns-repeated="1023"/></table:table-row><table:table-row table:style-name="ro1" table:number-rows-repeated="1048574"><table:table-cell table:number-columns-repeated="1024"/></table:table-row><table:table-row table:style-name="ro1"><table:table-cell table:number-columns-repeated="1024"/></table:table-row></table:table>`;
    const sheet = readOds(ods(body, { automaticStyles: optimal })).activeSheet;
    expect(sheet.rows.size).toBe(0);
    expect(sheet.getValue(0, 0)).toEqual(CellValue.number(1));
  });

  it('custom-height padding split into a run and a last row is skipped too', () => {
    const body = `<table:table table:name="S"><table:table-column/><table:table-row><table:table-cell office:value-type="float" office:value="1"/></table:table-row><table:table-row table:style-name="ro2" table:number-rows-repeated="1048574"><table:table-cell/></table:table-row><table:table-row table:style-name="ro2"><table:table-cell/></table:table-row></table:table>`;
    expect(readOds(ods(body, { automaticStyles: custom })).activeSheet.rows.size).toBe(0);
  });

  it('a set row height is written as not optimal, and an optimal one is not read as set', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.getOrCreateRow(0).height = 40;
    workbook.activeSheet.setValue(0, 0, CellValue.number(1));
    const saved = content(writeOds(workbook));
    expect(saved).toMatch(/style:row-height="[^"]+" style:use-optimal-row-height="false"/);
    expect(roundTrip(workbook).activeSheet.getRowHeight(0)).toBeCloseTo(40, 0);

    const body = `<table:table table:name="S"><table:table-column/><table:table-row table:style-name="ro1"><table:table-cell office:value-type="float" office:value="1"/></table:table-row></table:table>`;
    expect(readOds(ods(body, { automaticStyles: optimal })).activeSheet.rows.size).toBe(0);
  });
});

describe('generated style names', () => {
  it('never reuse a name styles.xml declares', () => {
    const styles = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS} office:version="1.3"><office:styles><number:number-style style:name="N0"><number:number number:decimal-places="0"/></number:number-style></office:styles><office:automatic-styles><number:date-style style:name="N1"><number:year/></number:date-style><number:number-style style:name="N2P0"><number:number/></number:number-style><style:style style:name="ce1" style:family="table-cell"/></office:automatic-styles></office:document-styles>`;
    const body = `<table:table table:name="S"><table:table-column/><table:table-row><table:table-cell office:value-type="float" office:value="1"/></table:table-row></table:table>`;
    const { workbook, package: pkg } = readOdsWithPackage(ods(body, { styles }));
    const sheet = workbook.activeSheet;
    sheet.getOrCreate(0, 0).styleId = workbook.styles.intern({
      ...CellStyles.default,
      numberFormat: NumberFormats.of('0.00;[Red]-0.00'),
    });
    const saved = content(writeOds(workbook, pkg));
    const names = [...saved.matchAll(/style:name="([^"]+)"/g)].map((m) => m[1]);
    expect(names.length).toBeGreaterThan(0);
    for (const taken of ['N0', 'N1', 'N2', 'N2P0', 'ce1']) expect(names).not.toContain(taken);
  });
});

describe('merges', () => {
  it('neighbouring merges are never folded into one repeated row or cell', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    for (let r = 0; r < 3; r++) sheet.addMerge(new MergedRegion(CellRange.fromBounds(r, 0, r, 1)));
    sheet.addMerge(new MergedRegion(CellRange.fromBounds(4, 0, 5, 0)));
    sheet.addMerge(new MergedRegion(CellRange.fromBounds(4, 1, 5, 1)));
    sheet.setValue(7, 3, CellValue.number(1));

    const saved = content(writeOds(workbook));
    expect(saved.match(/number-columns-spanned="2"/g)).toHaveLength(3);
    expect(saved.match(/number-rows-spanned="2"/g)).toHaveLength(2);
    expect(
      roundTrip(workbook)
        .activeSheet.merges.map((m) => m.range.toString())
        .sort(),
    ).toEqual(['A1:B1', 'A2:B2', 'A3:B3', 'A5:A6', 'B5:B6']);
  });

  it('an overlapping merge in a damaged file is skipped, not a failed load', () => {
    const body = `<table:table table:name="S"><table:table-column table:number-columns-repeated="3"/><table:table-row><table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="2" office:value-type="float" office:value="1"/><table:covered-table-cell/><table:table-cell/></table:table-row><table:table-row><table:covered-table-cell/><table:table-cell table:number-columns-spanned="2" office:value-type="float" office:value="2"/><table:table-cell/></table:table-row></table:table>`;
    const sheet = readOds(ods(body)).activeSheet;
    expect(sheet.merges.map((m) => m.range.toString())).toEqual(['A1:B2']);
    expect(sheet.getValue(1, 1)).toEqual(CellValue.number(2));
  });
});

describe('data validation messages and flags', () => {
  const body = `<table:content-validations><table:content-validation table:name="val1" table:condition="of:cell-content-is-whole-number() and cell-content()&gt;5" table:allow-empty-cell="false" table:base-cell-address="S.A1"><table:help-message table:title="Tip" table:display="true"><text:p>Enter<text:s text:c="2"/>a<text:line-break/>number</text:p></table:help-message><table:error-message table:message-type="warning" table:display="false" table:title="Oops"><text:p>bad</text:p></table:error-message></table:content-validation></table:content-validations><table:table table:name="S"><table:table-column/><table:table-row><table:table-cell table:content-validation-name="val1"/></table:table-row></table:table>`;

  it('reads a message with its spaces and line breaks', () => {
    const validation = readOds(ods(body)).activeSheet.validations[0]!;
    expect(validation.inputMessage).toBe('Enter  a\nnumber');
    expect(validation.errorMessage).toBe('bad');
    // display="false": LibreOffice lets a breaking entry through without the warning.
    expect(validation.response).toBe('none');
  });

  it('keeps allow-empty-cell, titles and display flags through a save', () => {
    const saved = content(resave(ods(body)));
    expect(saved).toContain('table:allow-empty-cell="false"');
    expect(saved).toMatch(/<table:help-message[^>]*table:title="Tip"/);
    expect(saved).toMatch(/<table:error-message[^>]*table:display="false"[^>]*table:title="Oops"/);
    expect(saved).toContain('table:message-type="warning"');
    expect(saved).toContain('<text:p>Enter <text:s/>a</text:p><text:p>number</text:p>');
    expect(readOds(resave(ods(body))).activeSheet.validations[0]!.inputMessage).toBe('Enter  a\nnumber');
  });

  it('writes a multi-line message as one paragraph per line', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.addValidation(
      dataValidation('wholeNumber', CellRange.parse('A1'), {
        operator: 'greaterThan',
        operand1: 0,
        errorMessage: 'Too small.\nTry again.',
      }),
    );
    expect(content(writeOds(workbook))).toContain('<text:p>Too small.</text:p><text:p>Try again.</text:p>');
  });
});

describe('cell values', () => {
  it('a cell with text but no value type is text', () => {
    const body = `<table:table table:name="S"><table:table-column/><table:table-row><table:table-cell><text:p>hello</text:p></table:table-cell></table:table-row></table:table>`;
    expect(readOds(ods(body)).activeSheet.getValue(0, 0)).toEqual(CellValue.text('hello'));
  });

  it('office:string-value is the value of a string cell', () => {
    const body = `<table:table table:name="S"><table:table-column/><table:table-row><table:table-cell office:value-type="string" office:string-value="hi"><text:p>shown</text:p></table:table-cell><table:table-cell office:value-type="string" office:string-value="only"/></table:table-row></table:table>`;
    const sheet = readOds(ods(body)).activeSheet;
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('hi'));
    expect(sheet.getValue(0, 1)).toEqual(CellValue.text('only'));
  });
});

describe('number formats', () => {
  function formatAfterRoundTrip(code: string): string {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.number(1234));
    sheet.getOrCreate(0, 0).styleId = workbook.styles.intern({
      ...CellStyles.default,
      numberFormat: NumberFormats.of(code),
    });
    const reopened = roundTrip(workbook);
    return reopened.styles.get(reopened.activeSheet.find(0, 0)!.styleId).numberFormat.code;
  }

  it.each(['00000', '#,###', '#', '0.0#', '#,##0.00', '00,000', '0.00'])('%s round-trips', (code) => {
    expect(formatAfterRoundTrip(code)).toBe(code);
  });

  it('reads LibreOffice min-integer-digits and min-decimal-places', () => {
    const automaticStyles = `<number:number-style style:name="N1"><number:number number:decimal-places="2" number:min-decimal-places="1" number:min-integer-digits="5"/></number:number-style><style:style style:name="ce1" style:family="table-cell" style:data-style-name="N1"/>`;
    const body = `<table:table table:name="S"><table:table-column/><table:table-row><table:table-cell table:style-name="ce1" office:value-type="float" office:value="1234"/></table:table-row></table:table>`;
    const workbook = readOds(ods(body, { automaticStyles }));
    expect(workbook.styles.get(workbook.activeSheet.find(0, 0)!.styleId).numberFormat.code).toBe('00000.0#');
  });
});

describe('database ranges', () => {
  it('a sorted range without filter buttons does not gain them, and its unmodelled parts stay', () => {
    const rows = [1, 2, 3]
      .map(
        (n) =>
          `<table:table-row><table:table-cell office:value-type="float" office:value="${n}"/></table:table-row>`,
      )
      .join('');
    const body = `<table:table table:name="S"><table:table-column/>${rows}</table:table><table:database-ranges><table:database-range table:name="__Anonymous_Sheet_DB__0" table:target-range-address="S.A1:S.A3" table:contains-header="true" table:refresh-delay="5"><table:sort table:case-sensitive="true"><table:sort-by table:field-number="0" table:data-type="number" table:order="ascending"/></table:sort><table:subtotal-rules/></table:database-range></table:database-ranges>`;
    const saved = content(resave(ods(body)));
    const range = saved.slice(
      saved.indexOf('<table:database-range '),
      saved.indexOf('</table:database-range>'),
    );
    expect(range).toContain('table:display-filter-buttons="false"');
    expect(range).toContain('table:refresh-delay="5"');
    expect(range).toContain('table:case-sensitive="true"');
    expect(range).toContain('table:data-type="number"');
    expect(range).toContain('<table:subtotal-rules/>');
  });

  it('an AutoFilter range keeps its buttons', () => {
    const body = `<table:table table:name="S"><table:table-column/><table:table-row><table:table-cell office:value-type="float" office:value="1"/></table:table-row></table:table><table:database-ranges><table:database-range table:name="__Anonymous_Sheet_DB__0" table:target-range-address="S.A1:S.A3" table:display-filter-buttons="true"/></table:database-ranges>`;
    expect(content(resave(ods(body)))).toContain('table:display-filter-buttons="true"');
  });
});

describe('checkboxes', () => {
  const control = `<draw:control table:end-cell-address="S.B2" table:end-x="1cm" table:end-y="0.4cm" draw:z-index="0" svg:width="1cm" svg:height="0.4cm" svg:x="0cm" svg:y="0cm" draw:control="control1"/>`;
  const forms = `<office:forms form:automatic-focus="false" form:apply-design-mode="false"><form:form form:name="Form"><form:checkbox form:name="Check Box 1" xml:id="control1" form:id="control1" form:linked-cell="S.B2" form:current-state="unchecked"/></form:form></office:forms>`;

  it('a checkbox linked to a cell that is not TRUE/FALSE is kept as it was', () => {
    const body = `<table:table table:name="S">${forms}<table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell>${control}</table:table-cell><table:table-cell office:value-type="float" office:value="3"/></table:table-row></table:table>`;
    const bytes = ods(body);
    const workbook = readOds(bytes);
    expect(workbook.styles.get(workbook.activeSheet.effectiveStyleId(1, 1)).checkbox ?? false).toBe(false);
    const saved = content(resave(bytes));
    expect(saved).toContain('form:linked-cell="S.B2"');
    expect(saved).toContain('draw:control="control1"');
  });

  it('a checkbox linked to a boolean cell is still taken over', () => {
    const body = `<table:table table:name="S">${forms}<table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell/></table:table-row><table:table-row><table:table-cell>${control}</table:table-cell><table:table-cell office:value-type="boolean" office:boolean-value="false"/></table:table-row></table:table>`;
    const workbook = readOds(ods(body));
    expect(workbook.styles.get(workbook.activeSheet.effectiveStyleId(1, 1)).checkbox).toBe(true);
  });
});

describe('settings.xml', () => {
  it('keeps the view and configuration settings the model does not own', () => {
    const settings = `<?xml version="1.0" encoding="UTF-8"?><office:document-settings ${NS} office:version="1.3"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item config:name="VisibleAreaWidth" config:type="int">9000</config:config-item><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry><config:config-item config:name="ViewId" config:type="string">view1</config:config-item><config:config-item-map-named config:name="Tables"><config:config-item-map-entry config:name="S"><config:config-item config:name="ZoomValue" config:type="int">150</config:config-item><config:config-item config:name="HorizontalSplitMode" config:type="short">0</config:config-item><config:config-item config:name="VerticalSplitMode" config:type="short">0</config:config-item></config:config-item-map-entry><config:config-item-map-entry config:name="Gone"/></config:config-item-map-named><config:config-item config:name="ActiveTable" config:type="string">S</config:config-item><config:config-item config:name="ShowGrid" config:type="boolean">false</config:config-item></config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set><config:config-item-set config:name="ooo:configuration-settings"><config:config-item config:name="PrinterName" config:type="string">Office</config:config-item></config:config-item-set></office:settings></office:document-settings>`;
    const body = `<table:table table:name="S"><table:table-column/><table:table-row><table:table-cell office:value-type="float" office:value="1"/></table:table-row></table:table>`;
    const { workbook, package: pkg } = readOdsWithPackage(ods(body, { settings }));
    workbook.activeSheet.frozenRows = 1;
    const saved = readEntry(writeOds(workbook, pkg), OdsXmlNames.settingsPart);

    expect(saved).toContain('config:name="ShowGrid" config:type="boolean">false<');
    expect(saved).toContain('config:name="ZoomValue" config:type="int">150<');
    expect(saved).toContain('config:name="VisibleAreaWidth"');
    expect(saved).toContain('config:name="PrinterName" config:type="string">Office<');
    expect(saved).toContain('config:name="VerticalSplitMode" config:type="short">2<');
    expect(saved).not.toContain('config:name="Gone"');
    expect(readOds(writeOds(workbook, pkg)).activeSheet.frozenRows).toBe(1);
  });
});

describe('CSV numbers', () => {
  const options = { ...CsvOptions.default, locale: 'en-US' };

  it('a group separator counts only where groups fall', () => {
    const sheet = readCsv('"1,2","1,234","12,345,678","10,20,30",1234\n', options).activeSheet;
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('1,2'));
    expect(sheet.getValue(0, 1)).toEqual(CellValue.number(1234));
    expect(sheet.getValue(0, 2)).toEqual(CellValue.number(12345678));
    expect(sheet.getValue(0, 3)).toEqual(CellValue.text('10,20,30'));
    expect(sheet.getValue(0, 4)).toEqual(CellValue.number(1234));
  });
});
