import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { Defaults } from '../../src/defaults.ts';
import { namedRangeDefinition } from '../../src/model/namedRange.ts';
import { describeXlsxImport, OdsUnsupportedFileError, XlsxReader } from '../../src/serialization/index.ts';
import { applyTint } from '../../src/serialization/xlsx/xlsxColors.ts';
import { excelToInternal } from '../../src/serialization/xlsx/xlsxFormulas.ts';
import { Argb } from '../../src/styling/color.ts';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';

interface Book {
  sheets?: { name: string; xml: string; state?: string; rels?: string }[];
  workbookExtra?: string;
  workbookPr?: string;
  styles?: string;
  strings?: string[];
  parts?: Record<string, string>;
}

const rel = (id: string, type: string, target: string, external = false) =>
  `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;

/** A minimal `.xlsx` package from hand-written parts. */
function xlsx(book: Book): Uint8Array {
  const sheets = book.sheets ?? [{ name: 'Sheet1', xml: '<sheetData/>' }];
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${PR}">${rel('r1', 'officeDocument', 'xl/workbook.xml')}</Relationships>`,
    ),
    'xl/workbook.xml': strToU8(
      `<workbook xmlns="${MAIN}" xmlns:r="${R}">${book.workbookPr ?? ''}<sheets>${sheets
        .map(
          (s, i) =>
            `<sheet name="${s.name}" sheetId="${i + 1}" r:id="s${i}"${s.state ? ` state="${s.state}"` : ''}/>`,
        )
        .join('')}</sheets>${book.workbookExtra ?? ''}</workbook>`,
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `<Relationships xmlns="${PR}">${sheets.map((_, i) => rel(`s${i}`, 'worksheet', `worksheets/sheet${i + 1}.xml`)).join('')}${
        book.styles ? rel('st', 'styles', 'styles.xml') : ''
      }${book.strings ? rel('ss', 'sharedStrings', 'sharedStrings.xml') : ''}</Relationships>`,
    ),
  };
  sheets.forEach((s, i) => {
    files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(
      `<worksheet xmlns="${MAIN}" xmlns:r="${R}">${s.xml}</worksheet>`,
    );
    if (s.rels)
      files[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] = strToU8(
        `<Relationships xmlns="${PR}">${s.rels}</Relationships>`,
      );
  });
  if (book.styles)
    files['xl/styles.xml'] = strToU8(`<styleSheet xmlns="${MAIN}">${book.styles}</styleSheet>`);
  if (book.strings)
    files['xl/sharedStrings.xml'] = strToU8(
      `<sst xmlns="${MAIN}">${book.strings.map((s) => `<si>${s}</si>`).join('')}</sst>`,
    );
  for (const [path, text] of Object.entries(book.parts ?? {})) files[path] = strToU8(text);
  return zipSync(files);
}

const read = (book: Book) => XlsxReader.read(xlsx(book), 'en-US');
const sheetOf = (cells: string) => ({ name: 'Sheet1', xml: `<sheetData>${cells}</sheetData>` });

describe('Excel formulas in Cellestra syntax', () => {
  it.each([
    ['SUM(A1:B2)', 'SUM(A1:B2)'],
    ['Sheet2!A1+1', 'Sheet2.A1+1'],
    ["'My Sheet'!$A$1", "'My Sheet'.$A$1"],
    ["'It''s'!B2", "'It''s'.B2"],
    ['SUM(Jan:Mar!B2:C3)', 'SUM(Jan.B2:Mar.C3)'],
    ['_xlfn.CONCAT("a!b",A1)', 'CONCAT("a!b",A1)'],
    ['_xlfn._xlws.SORT(A1:A3)', 'SORT(A1:A3)'],
  ])('%s → %s', (excel, internal) => {
    expect(excelToInternal(excel)).toBe(internal);
  });

  it('has no syntax for tables or other workbooks', () => {
    expect(excelToInternal('SUM(Table1[Total])')).toBeNull();
    expect(excelToInternal('[1]Sheet1!A1')).toBeNull();
  });
});

describe('reading an .xlsx', () => {
  it('reads every value type', () => {
    const { workbook } = read({
      strings: [
        '<t>hello</t>',
        '<r><t>rich </t></r><r><t>text</t></r><rPh><t>skip</t></rPh>',
        '<t>line_x000D_</t>',
      ],
      sheets: [
        sheetOf(
          '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1"><v>2.5</v></c>' +
            '<c r="D1" t="b"><v>1</v></c><c r="E1" t="e"><v>#N/A</v></c><c r="F1" t="inlineStr"><is><t>inline</t></is></c>' +
            '<c r="G1" t="s"><v>2</v></c><c r="H1" t="d"><v>2026-03-01T12:00:00</v></c></row>',
        ),
      ],
    });
    const s = workbook.worksheets[0]!;
    expect(s.getValue(0, 0)).toEqual({ kind: 'text', value: 'hello' });
    expect(s.getValue(0, 1)).toEqual({ kind: 'text', value: 'rich text' });
    expect(s.getValue(0, 2)).toEqual({ kind: 'number', value: 2.5 });
    expect(s.getValue(0, 3)).toEqual({ kind: 'boolean', value: true });
    expect(s.getValue(0, 4).kind).toBe('error');
    expect(s.getValue(0, 5)).toEqual({ kind: 'text', value: 'inline' });
    expect(s.getValue(0, 6)).toEqual({ kind: 'text', value: 'line\r' });
    expect(s.getValue(0, 7)).toEqual({ kind: 'number', value: 46082.5 });
  });

  it('places cells without an r after the one before', () => {
    const { workbook } = read({
      sheets: [sheetOf('<row><c><v>1</v></c><c><v>2</v></c></row><row><c><v>3</v></c></row>')],
    });
    const s = workbook.worksheets[0]!;
    expect([s.getValue(0, 1), s.getValue(1, 0)]).toEqual([
      { kind: 'number', value: 2 },
      { kind: 'number', value: 3 },
    ]);
  });

  it('keeps formulas with their cached results and expands shared ones', () => {
    const { workbook, report } = read({
      sheets: [
        sheetOf(
          '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row>' +
            '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>' +
            '<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c>' +
            '<c r="C3"><f>NOSUCHFUNCTION(A1)</f><v>7</v></c><c r="D3"><f t="array" ref="D3">SUM(A1:A3*2)</f><v>12</v></c>' +
            '<c r="E3"><f>Missing*2</f><v>8</v></c></row>',
        ),
      ],
    });
    const s = workbook.worksheets[0]!;
    expect([s.find(0, 1)?.formula, s.find(1, 1)?.formula, s.find(2, 1)?.formula]).toEqual([
      'A1*2',
      'A2*2',
      'A3*2',
    ]);
    expect(s.find(2, 1)?.value).toEqual({ kind: 'number', value: 6 });
    // Unknown functions, array formulas and unknown names keep the file's result.
    expect([s.find(2, 2)?.formula, s.find(2, 3)?.formula, s.find(2, 4)?.formula]).toEqual([null, null, null]);
    expect(s.getValue(2, 2)).toEqual({ kind: 'number', value: 7 });
    expect(report.formulasAsValues).toBe(3);
  });

  it('reads named ranges, global and per sheet', () => {
    const { workbook, report } = read({
      sheets: [
        sheetOf('<row r="1"><c r="A1"><f>Rate*2</f><v>0.2</v></c></row>'),
        { name: 'Two', xml: '<sheetData/>' },
      ],
      workbookExtra:
        '<definedNames><definedName name="Rate">Sheet1!$B$1</definedName>' +
        '<definedName name="Local" localSheetId="1">Two!$A$1:$A$3</definedName>' +
        '<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$C$3</definedName>' +
        '<definedName name="Constant">0.07</definedName>' +
        '<definedName name="Broken">SUM(</definedName></definedNames>',
    });
    expect(workbook.namedRanges.map((n) => [n.name, namedRangeDefinition(n), n.scope ?? null])).toEqual([
      ['Rate', 'Sheet1.$B$1', null],
      ['Local', 'Two.$A$1:$A$3', 'Two'],
      ['Constant', '=0.07', null],
    ]);
    expect(workbook.worksheets[0]!.find(0, 0)?.formula).toBe('Rate*2');
    expect(report.leftOut.has('names')).toBe(true);
  });

  it('reads widths, heights, hidden lines, merges and frozen panes', () => {
    const { workbook } = read({
      sheets: [
        {
          name: 'Sheet1',
          xml:
            '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" state="frozen"/></sheetView></sheetViews>' +
            '<sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="9.140625"/><col min="2" max="3" width="20" customWidth="1"/>' +
            '<col min="4" max="4" width="9" hidden="1"/></cols>' +
            '<sheetData><row r="1" ht="30" customHeight="1"><c r="A1"><v>1</v></c></row><row r="2" ht="15"/><row r="3" hidden="1"/></sheetData>' +
            '<mergeCells count="1"><mergeCell ref="A5:C6"/></mergeCells>',
        },
      ],
    });
    const s = workbook.worksheets[0]!;
    expect([s.getColumnWidth(0), s.getColumnWidth(1), s.getColumnWidth(2)]).toEqual([
      Defaults.columnWidth,
      140,
      140,
    ]);
    expect(s.isColumnHidden(3)).toBe(true);
    expect([s.getRowHeight(0), s.getRowHeight(1)]).toEqual([40, Defaults.rowHeight]);
    expect(s.isRowHidden(2)).toBe(true);
    expect(s.merges.map((m) => m.range.toString())).toEqual(['A5:C6']);
    expect([s.frozenRows, s.frozenColumns]).toEqual([2, 1]);
  });

  it('reads cell styles: fonts, theme colours with tint, fills, borders, alignment and number formats', () => {
    const { workbook } = read({
      styles:
        '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>' +
        '<fonts><font><sz val="11"/><name val="Calibri"/></font><font><b/><i/><u/><sz val="14"/><color rgb="FFFF0000"/><name val="Georgia"/></font></fonts>' +
        '<fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor theme="4" tint="0.5"/></patternFill></fill></fills>' +
        '<borders><border/><border><left style="thin"><color indexed="10"/></left><bottom style="double"/></border></borders>' +
        '<cellXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' +
        '<xf numFmtId="164" fontId="1" fillId="2" borderId="1"><alignment horizontal="center" vertical="center" wrapText="1" textRotation="135"/></xf>' +
        '<xf numFmtId="10" fontId="0" fillId="0" borderId="0"/></cellXfs>',
      sheets: [
        sheetOf(
          '<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>0.5</v></c><c r="C1"><v>1</v></c></row>',
        ),
      ],
    });
    const s = workbook.worksheets[0]!;
    const a1 = workbook.styles.get(s.find(0, 0)!.styleId);
    expect(a1.numberFormat.code).toBe('0.000');
    expect(a1.font).toMatchObject({
      family: 'Georgia',
      sizePoints: 14,
      bold: true,
      italic: true,
      underline: true,
    });
    expect(a1.font.color).toBe(Argb.fromRgb(255, 0, 0));
    expect(a1.fill).toEqual({ background: applyTint(0x4472c4, 0.5) });
    expect(a1.border.left).toMatchObject({ style: 'thin', color: Argb.fromRgb(255, 0, 0) });
    expect(a1.border.bottom).toMatchObject({ style: 'double' });
    expect(a1.alignment).toEqual({ horizontal: 'center', vertical: 'middle', wrapText: true, rotation: 315 });
    expect(workbook.styles.get(s.find(0, 1)!.styleId).numberFormat.code).toBe('0.00%');
    // The workbook's default font is Cellestra's default: plain cells carry no style.
    expect(s.find(0, 2)!.styleId).toBe(0);
  });

  it('lightens and darkens by tint as Excel does', () => {
    expect(applyTint(0x000000, 0.5)).toBe(Argb.fromRgb(128, 128, 128));
    expect(applyTint(0xffffff, -0.5)).toBe(Argb.fromRgb(128, 128, 128));
  });

  it('shifts dates of the 1904 date system to Cellestra’s', () => {
    const { workbook } = read({
      workbookPr: '<workbookPr date1904="1"/>',
      styles: '<cellXfs><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs>',
      sheets: [sheetOf('<row r="1"><c r="A1" s="1"><v>0</v></c><c r="B1"><v>5</v></c></row>')],
    });
    const s = workbook.worksheets[0]!;
    expect(s.getValue(0, 0)).toEqual({ kind: 'number', value: 1462 });
    expect(s.getValue(0, 1)).toEqual({ kind: 'number', value: 5 });
  });

  it('reads web and in-workbook hyperlinks', () => {
    const { workbook } = read({
      sheets: [
        {
          name: 'Sheet1',
          xml:
            '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>site</t></is></c></row></sheetData>' +
            '<hyperlinks><hyperlink ref="A1" r:id="h1" tooltip="Go"/><hyperlink ref="B2" location="\'Other Sheet\'!C3"/></hyperlinks>',
          rels: rel('h1', 'hyperlink', 'https://example.com/', true),
        },
      ],
    });
    const s = workbook.worksheets[0]!;
    expect(s.find(0, 0)?.hyperlink).toEqual({ target: 'https://example.com/', tooltip: 'Go' });
    expect(s.find(1, 1)?.hyperlink).toEqual({ target: "#'Other Sheet'.C3" });
  });

  it('opens hidden sheets visible, picks the active one, and reports what it left out', () => {
    const { workbook, report } = read({
      workbookExtra: '<bookViews><workbookView activeTab="1"/></bookViews>',
      sheets: [
        {
          name: 'A',
          xml: '<sheetData/><conditionalFormatting sqref="A1"/><dataValidations><dataValidation type="whole" sqref="A1"><formula1>1</formula1></dataValidation></dataValidations><autoFilter ref="A1:B2"/><drawing r:id="d"/>',
        },
        { name: 'B', xml: '<sheetData/>', state: 'hidden' },
      ],
    });
    expect(workbook.worksheets.map((s) => s.name)).toEqual(['A', 'B']);
    expect(workbook.activeSheet.name).toBe('B');
    expect(describeXlsxImport(report)).toBe(
      'Left out: charts, images and shapes, conditional formatting, data validation, AutoFilters. 1 hidden sheet is shown.',
    );
  });

  it('reads dropdown lists: typed, from a range on another sheet, by name, and in the x14 form', () => {
    const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
    const XM = 'http://schemas.microsoft.com/office/excel/2006/main';
    const { workbook, report } = read({
      workbookExtra: '<definedNames><definedName name="Sizes">Options!$B$1:$B$3</definedName></definedNames>',
      sheets: [
        {
          name: 'Sheet1',
          xml:
            '<sheetData/><dataValidations count="4">' +
            '<dataValidation type="list" allowBlank="1" sqref="A1:A5 C1"><formula1>"Yes, No,""Maybe"""</formula1></dataValidation>' +
            `<dataValidation type="list" showErrorMessage="1" errorStyle="warning" error="Pick one" showInputMessage="1" prompt="Choose" sqref="B2:B9"><formula1>'Dropdown options'!$A$1:$A$6</formula1></dataValidation>` +
            '<dataValidation type="list" showErrorMessage="1" sqref="D1"><formula1>Sizes</formula1></dataValidation>' +
            '<dataValidation type="list" sqref="E1"><formula1>INDIRECT(A1)</formula1></dataValidation>' +
            '</dataValidations>' +
            `<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations xmlns:x14="${X14}" xmlns:xm="${XM}">` +
            '<x14:dataValidation type="list" showErrorMessage="1"><x14:formula1><xm:f>Options!$A$1:$A$2</xm:f></x14:formula1><xm:sqref>F1:F3</xm:sqref></x14:dataValidation>' +
            '</x14:dataValidations></ext></extLst>',
        },
        { name: 'Dropdown options', xml: '<sheetData/>' },
        { name: 'Options', xml: '<sheetData/>' },
      ],
    });
    const rules = workbook.worksheets[0]!.validations.map((v) => ({
      target: v.target.toString(),
      values: v.listValues ?? null,
      source: v.listSource?.toString() ?? null,
      response: v.response,
      input: v.inputMessage ?? null,
      error: v.errorMessage ?? null,
    }));
    const typed = {
      values: ['Yes', 'No', '"Maybe"'],
      source: null,
      response: 'none',
      input: null,
      error: null,
    };
    expect(rules).toEqual([
      { target: 'A1:A5', ...typed },
      { target: 'C1', ...typed },
      {
        target: 'B2:B9',
        values: null,
        source: "'Dropdown options'.$A$1:$A$6",
        response: 'warning',
        input: 'Choose',
        error: 'Pick one',
      },
      { target: 'D1', values: null, source: 'Options.$B$1:$B$3', response: 'stop', input: null, error: null },
      {
        target: 'F1:F3',
        values: null,
        source: 'Options.$A$1:$A$2',
        response: 'stop',
        input: null,
        error: null,
      },
    ]);
    // A list from a formula such as INDIRECT can't come across.
    expect(report.leftOut.has('dataValidation')).toBe(true);
  });

  it('says nothing when everything came across', () => {
    expect(describeXlsxImport(read({}).report)).toBe('');
  });

  it('refuses encrypted workbooks, .xls files and other bytes with a message for the person', () => {
    const ole = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0]);
    expect(() => XlsxReader.read(ole)).toThrow(OdsUnsupportedFileError);
    expect(() => XlsxReader.read(ole)).toThrow(/password-protected or in the old \.xls format/);
    expect(() => XlsxReader.read(strToU8('not a zip'))).toThrow(/not an \.xlsx workbook/);
  });
});
