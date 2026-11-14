import { createChart, createSeries } from '@cellestra/charts';
import { describe, expect, it } from 'vitest';
import { dataValidation } from '../../src/data/types.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { MergedRegion } from '../../src/model/mergedRegion.ts';
import { SheetDimensions } from '../../src/model/sheetDimensions.ts';
import { Workbook } from '../../src/model/workbook.ts';
import {
  CsvOptions,
  OdsPackage,
  OdsPackageTooLargeError,
  OdsXmlNames,
  readOds,
  readOdsWithPackage,
  writeCsv,
  writeOds,
} from '../../src/serialization/index.ts';
import { CellStyles, NumberFormats } from '../../src/styling/cellStyle.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { OdsContentTooLargeError } from '../../src/serialization/ods/odsContentRead.ts';
import { entryNames, readEntry, roundTrip } from './helpers.ts';

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
  'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"',
].join(' ');

interface PackageParts {
  readonly automaticStyles?: string;
  readonly fontFaces?: string;
  readonly beforeTables?: string;
  readonly afterTables?: string;
  readonly styles?: string;
  readonly extra?: Record<string, string>;
  readonly manifest?: string;
}

/** A hand-built `.ods` whose content.xml body is `tables`. */
function ods(tables: string, parts: PackageParts = {}): Uint8Array {
  const content = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${NS} office:version="1.3"><office:font-face-decls>${parts.fontFaces ?? ''}</office:font-face-decls><office:automatic-styles>${parts.automaticStyles ?? ''}</office:automatic-styles><office:body><office:spreadsheet>${parts.beforeTables ?? ''}${tables}${parts.afterTables ?? ''}</office:spreadsheet></office:body></office:document-content>`;
  const pkg = new OdsPackage();
  pkg.setText(OdsXmlNames.mimetypePart, OdsXmlNames.mimeType);
  pkg.setText(OdsXmlNames.contentPart, content);
  if (parts.styles !== undefined) pkg.setText(OdsXmlNames.stylesPart, parts.styles);
  if (parts.manifest !== undefined) pkg.setText(OdsXmlNames.manifestPart, parts.manifest);
  for (const [name, text] of Object.entries(parts.extra ?? {})) pkg.setText(name, text);
  return pkg.write();
}

const float = (n: number): string =>
  `<table:table-cell office:value-type="float" office:value="${n}"><text:p>${n}</text:p></table:table-cell>`;

/** Open, save through the model with the source package, and return the saved bytes. */
function resave(bytes: Uint8Array): Uint8Array {
  const { workbook, package: pkg } = readOdsWithPackage(bytes);
  return writeOds(workbook, pkg);
}

function formatAfterRoundTrip(code: string, value = 0.5): string {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  sheet.setValue(0, 0, CellValue.number(value));
  sheet.getOrCreate(0, 0).styleId = workbook.styles.intern({
    ...CellStyles.default,
    numberFormat: NumberFormats.of(code),
  });
  const reopened = roundTrip(workbook);
  return reopened.styles.get(reopened.activeSheet.find(0, 0)!.styleId).numberFormat.code;
}

function contentOf(workbook: Workbook): string {
  return readEntry(writeOds(workbook), OdsXmlNames.contentPart);
}

describe('ODS grid structure', () => {
  it('reads rows and columns inside header, group and rows containers with one running index', () => {
    const bytes = ods(
      `<table:table table:name="Sheet1">
        <table:table-header-columns><table:table-column table:style-name="co1"/></table:table-header-columns>
        <table:table-column-group><table:table-column table:style-name="co2"/></table:table-column-group>
        <table:table-header-rows><table:table-row>${float(1)}</table:table-row></table:table-header-rows>
        <table:table-row-group>
          <table:table-row-group><table:table-row>${float(2)}</table:table-row></table:table-row-group>
          <table:table-rows><table:table-row>${float(3)}</table:table-row></table:table-rows>
        </table:table-row-group>
        <table:table-row>${float(4)}</table:table-row>
      </table:table>`,
      {
        automaticStyles: `<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="5cm"/></style:style>
          <style:style style:name="co2" style:family="table-column"><style:table-column-properties style:column-width="1in"/></style:style>`,
      },
    );
    const sheet = readOds(bytes).activeSheet;
    expect([0, 1, 2, 3].map((r) => sheet.getValue(r, 0))).toEqual(
      [1, 2, 3, 4].map((n) => CellValue.number(n)),
    );
    expect(sheet.getColumnWidth(0)).toBeGreaterThan(sheet.getColumnWidth(1));
    expect(sheet.getColumnWidth(1)).toBeGreaterThan(80);
  });

  it('keeps the height of blank rows between rows of data, but not of the padding after them', () => {
    const tall = `<table:table-row table:style-name="ro2"><table:table-cell/></table:table-row>`;
    const bytes = ods(
      `<table:table table:name="Sheet1">
        <table:table-row>${float(1)}</table:table-row>${tall}${tall}
        <table:table-row>${float(2)}</table:table-row>${tall}${tall}
        <table:table-row table:style-name="ro3"><table:table-cell/></table:table-row>
        <table:table-row>${float(3)}</table:table-row>
        <table:table-row table:style-name="ro2" table:number-rows-repeated="1000"><table:table-cell/></table:table-row>${tall}
      </table:table>`,
      {
        automaticStyles: `<style:style style:name="ro2" style:family="table-row"><style:table-row-properties style:row-height="2cm"/></style:style>
          <style:style style:name="ro3" style:family="table-row"><style:table-row-properties style:row-height="3cm"/></style:style>`,
      },
    );
    const sheet = readOds(bytes).activeSheet;
    const plain = sheet.getRowHeight(0);
    const heights = [1, 2, 4, 5, 6].map((r) => sheet.getRowHeight(r));
    expect(heights.every((h) => h > plain)).toBe(true);
    expect(heights[4]).toBeGreaterThan(heights[0]!);
    expect(sheet.getValue(7, 0)).toEqual(CellValue.number(3));
    expect([8, 500, 1008].map((r) => sheet.getRowHeight(r))).toEqual([plain, plain, plain]);
  });

  it('clamps hostile repeats and spans to the sheet size without hanging', () => {
    const bytes = ods(
      `<table:table table:name="Sheet1">
        <table:table-column table:number-columns-repeated="2000000000"/>
        <table:table-row table:number-rows-repeated="2000000000">
          <table:table-cell table:style-name="ce1" table:number-columns-repeated="2000000000"/>
        </table:table-row>
        <table:table-row><table:table-cell table:number-columns-spanned="99999999" table:number-rows-spanned="99999999"/></table:table-row>
      </table:table>`,
      {
        automaticStyles: `<style:style style:name="ce1" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style>`,
      },
    );
    const sheet = readOds(bytes).activeSheet;
    expect(sheet.cellCount).toBeLessThan(10);
    for (const merge of sheet.merges) {
      expect(merge.range.right).toBeLessThan(16384);
      expect(merge.range.bottom).toBeLessThan(1048576);
    }
  });

  describe('repeats are budgeted across the workbook', () => {
    // 1,024 rows of 2,049 value cells: two million cells from a few hundred bytes. Two of them
    // pass the workbook's budget of 4,194,304.
    const bigRow = `<table:table-row table:number-rows-repeated="1024"><table:table-cell table:number-columns-repeated="2049" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>`;

    it('refuses a sheet whose repeated rows add up past the budget', () => {
      const bytes = ods(`<table:table table:name="Sheet1">${bigRow}${bigRow}</table:table>`);
      expect(() => readOds(bytes)).toThrow(OdsContentTooLargeError);
    });

    it('refuses repeats spread over several sheets', () => {
      const sheet = (name: string): string => `<table:table table:name="${name}">${bigRow}</table:table>`;
      const bytes = ods(sheet('One') + sheet('Two'));
      expect(() => readOds(bytes)).toThrow(OdsContentTooLargeError);
    });

    it('refuses a spanned cell repeated into millions of merges', () => {
      const bytes = ods(
        `<table:table table:name="Sheet1"><table:table-row table:number-rows-repeated="1048576"><table:table-cell table:number-columns-spanned="2" table:number-columns-repeated="16384"/></table:table-row></table:table>`,
      );
      expect(() => readOds(bytes)).toThrow(OdsContentTooLargeError);
    });

    it('reads a repeated spanned cell as the merges that do not overlap', () => {
      // Each copy spans two rows, so the copies in odd rows would overlap the ones above them.
      const bytes = ods(
        `<table:table table:name="Sheet1"><table:table-row table:number-rows-repeated="6"><table:table-cell table:number-rows-spanned="2" table:number-columns-repeated="3"/></table:table-row></table:table>`,
      );
      const merges = readOds(bytes)
        .activeSheet.merges.map((m) => m.range.toString())
        .sort();
      expect(merges).toEqual(
        ['A1:A2', 'B1:B2', 'C1:C2', 'A3:A4', 'B3:B4', 'C3:C4', 'A5:A6', 'B5:B6', 'C5:C6'].sort(),
      );
    });
  });

  it('a styled run across a whole row becomes a style, not 16384 cells', () => {
    const bytes = ods(
      `<table:table table:name="Sheet1">
        <table:table-row><table:table-cell table:style-name="ce1" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell><table:table-cell table:style-name="ce1" table:number-columns-repeated="16383"/></table:table-row>
      </table:table>`,
      {
        automaticStyles: `<style:style style:name="ce1" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style>`,
      },
    );
    const workbook = readOds(bytes);
    const sheet = workbook.activeSheet;
    expect(sheet.cellCount).toBeLessThan(10);
    expect(workbook.styles.get(sheet.effectiveStyleId(0, 9000)).font.bold).toBe(true);
    expect(sheet.getValue(0, 0)).toEqual(CellValue.number(1));
  });

  it('a styled run down the whole column becomes a column style', () => {
    const bytes = ods(
      `<table:table table:name="Sheet1">
        <table:table-row table:number-rows-repeated="1048576"><table:table-cell table:style-name="ce1"/><table:table-cell/></table:table-row>
      </table:table>`,
      {
        automaticStyles: `<style:style style:name="ce1" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style>`,
      },
    );
    const workbook = readOds(bytes);
    const sheet = workbook.activeSheet;
    expect(sheet.cellCount).toBe(0);
    expect(workbook.styles.get(sheet.findColumn(0)!.styleId).font.bold).toBe(true);
  });

  describe('columns formatted below a header row', () => {
    // A header in A1:F1 over columns A:F formatted from row 2 to the end of the sheet, written as
    // LibreOffice writes it: one repeated row of a run per column (here each its own fill).
    const fills = ['ff0000', '00ff00', '0000ff', 'ffff00', '00ffff', 'ff00ff'];
    const automaticStyles =
      `<style:style style:name="ce1" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style>` +
      fills
        .map(
          (fill, i) =>
            `<style:style style:name="ce${i + 2}" style:family="table-cell"><style:table-cell-properties fo:background-color="#${fill}"/></style:style>`,
        )
        .join('');
    const headerOver = (headerStyle: string | null): Uint8Array => {
      const style = headerStyle === null ? '' : ` table:style-name="${headerStyle}"`;
      const header = ['A', 'B', 'C', 'D', 'E', 'F']
        .map(
          (h) =>
            `<table:table-cell${style} office:value-type="string"><text:p>${h}</text:p></table:table-cell>`,
        )
        .join('');
      const runs = fills.map((_, i) => `<table:table-cell table:style-name="ce${i + 2}"/>`).join('');
      return ods(
        `<table:table table:name="Sheet1">
          <table:table-row>${header}<table:table-cell table:number-columns-repeated="16378"/></table:table-row>
          <table:table-row table:number-rows-repeated="1048575">${runs}<table:table-cell table:number-columns-repeated="16378"/></table:table-row>
        </table:table>`,
        { automaticStyles },
      );
    };
    const fillAt = (workbook: Workbook, row: number, column: number): number | null =>
      workbook.styles.get(workbook.activeSheet.effectiveStyleId(row, column)).fill.background;
    const boldAt = (workbook: Workbook, row: number, column: number): boolean =>
      workbook.styles.get(workbook.activeSheet.effectiveStyleId(row, column)).font.bold;

    it('a styled header keeps its style and the columns become column formats', () => {
      // Held as cells, these six runs are 6,291,450 cells: past the repeat budget, so refused.
      const workbook = readOds(headerOver('ce1'));
      expect(workbook.activeSheet.cellCount).toBe(6);
      for (const reopened of [workbook, roundTrip(workbook)]) {
        expect(reopened.activeSheet.cellCount).toBe(6);
        expect(boldAt(reopened, 0, 0)).toBe(true);
        expect(fillAt(reopened, 0, 0)).toBeNull();
        expect(fillAt(reopened, 1, 0)! >>> 0).toBe(0xffff0000);
        expect(fillAt(reopened, SheetDimensions.maxRows - 1, 5)! >>> 0).toBe(0xffff00ff);
        expect(fillAt(reopened, 1, 6)).toBeNull();
      }
      const content = readEntry(writeOds(workbook), OdsXmlNames.contentPart);
      expect(content).toContain('table:default-cell-style-name');
    });

    it('a header in the default style stays plain over the formatted columns', () => {
      const workbook = readOds(headerOver(null));
      for (const reopened of [workbook, roundTrip(workbook)]) {
        expect(reopened.activeSheet.cellCount).toBe(6);
        expect(reopened.activeSheet.getValue(0, 2)).toEqual(CellValue.text('C'));
        expect(fillAt(reopened, 0, 2)).toBeNull();
        expect(boldAt(reopened, 0, 2)).toBe(false);
        expect(fillAt(reopened, 1, 2)! >>> 0).toBe(0xff0000ff);
      }
    });

    it('a blank row above the header is kept plain too', () => {
      const bytes = ods(
        `<table:table table:name="Sheet1">
          <table:table-row><table:table-cell table:number-columns-repeated="16384"/></table:table-row>
          <table:table-row><table:table-cell table:style-name="ce1" office:value-type="string"><text:p>A</text:p></table:table-cell><table:table-cell table:number-columns-repeated="16383"/></table:table-row>
          <table:table-row table:number-rows-repeated="1048574"><table:table-cell table:style-name="ce2"/><table:table-cell table:number-columns-repeated="16383"/></table:table-row>
        </table:table>`,
        { automaticStyles },
      );
      for (const workbook of [readOds(bytes), roundTrip(readOds(bytes))]) {
        expect(fillAt(workbook, 0, 0)).toBeNull();
        expect(boldAt(workbook, 1, 0)).toBe(true);
        expect(fillAt(workbook, 2, 0)! >>> 0).toBe(0xffff0000);
        expect(workbook.activeSheet.cellCount).toBe(2);
      }
    });
  });

  it('a styled run to the end of a row after a cell styled otherwise stays cells', () => {
    // A row format would restyle A1, which has the default style.
    const bytes = ods(
      `<table:table table:name="Sheet1">
        <table:table-row>${float(1)}<table:table-cell table:style-name="ce1" table:number-columns-repeated="16383"/></table:table-row>
      </table:table>`,
      {
        automaticStyles: `<style:style style:name="ce1" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style>`,
      },
    );
    const workbook = readOds(bytes);
    const sheet = workbook.activeSheet;
    expect(workbook.styles.get(sheet.effectiveStyleId(0, 9000)).font.bold).toBe(true);
    expect(workbook.styles.get(sheet.effectiveStyleId(0, 0)).font.bold).toBe(false);
  });

  it('reads deeply nested markup without recursion, and refuses absurd nesting quickly', () => {
    const cellWith = (depth: number): Uint8Array =>
      ods(
        `<table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="string"><text:p>${'<text:span>'.repeat(depth)}deep${'</text:span>'.repeat(depth)}</text:p></table:table-cell></table:table-row></table:table>`,
      );
    expect(readOds(cellWith(9_000)).activeSheet.getValue(0, 0)).toEqual(CellValue.text('deep'));
    const started = performance.now();
    expect(() => readOds(cellWith(200_000))).toThrow(/nested more than/);
    expect(performance.now() - started).toBeLessThan(3_000);
  });

  it('refuses a package whose declared uncompressed size is over the cap', () => {
    const bytes = ods(`<table:table table:name="Sheet1"/>`, {
      extra: { 'Pictures/big.txt': 'x'.repeat(10_000) },
    });
    expect(() => OdsPackage.read(bytes, 5_000)).toThrow(OdsPackageTooLargeError);
    expect(() => OdsPackage.read(bytes)).not.toThrow();
  });

  it('a normal sheet with an embedded object stays a normal sheet', () => {
    const bytes = ods(
      `<table:table table:name="Sheet1"><table:table-row>${float(7)}<table:table-cell><draw:frame svg:width="5cm" svg:height="5cm"><draw:object xlink:href="./Object 1"/></draw:frame></table:table-cell></table:table-row></table:table>`,
      {
        extra: {
          'Object 1/content.xml': `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS}><office:body><office:chart/></office:body></office:document-content>`,
        },
      },
    );
    const sheet = readOds(bytes).activeSheet;
    expect(sheet.isChartSheet).toBe(false);
    expect(sheet.getValue(0, 0)).toEqual(CellValue.number(7));
  });

  it('identical rows are written as one repeated row', () => {
    const workbook = Workbook.createNew();
    const bold = workbook.styles.intern({
      ...CellStyles.default,
      font: { ...CellStyles.default.font, bold: true },
    });
    for (let r = 0; r < 500; r++) workbook.activeSheet.getOrCreate(r, 1).styleId = bold;
    const content = contentOf(workbook);
    expect(content).toContain('table:number-rows-repeated="500"');
    expect(content.match(/<table:table-row\b/g)!.length).toBeLessThan(5);
  });

  it('covered cells are written as one repeated run and keep their content', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.addMerge(new MergedRegion(CellRange.parse('A1:Z1')));
    sheet.setValue(0, 5, CellValue.number(9)); // hidden under the merge
    const content = contentOf(workbook);
    expect(content).toMatch(/<table:covered-table-cell table:number-columns-repeated="4"\/>/);
    expect(content).toMatch(/<table:covered-table-cell table:number-columns-repeated="20"\/>/);
    expect(roundTrip(workbook).activeSheet.getValue(0, 5)).toEqual(CellValue.number(9));
  });

  it('reads the content of a covered cell', () => {
    const bytes = ods(
      `<table:table table:name="Sheet1"><table:table-row><table:table-cell table:number-columns-spanned="2" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell><table:covered-table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:covered-table-cell></table:table-row></table:table>`,
    );
    expect(readOds(bytes).activeSheet.getValue(0, 1)).toEqual(CellValue.number(2));
  });
});

describe('ODS unmodeled content', () => {
  const tables = `<table:table table:name="Sheet1" table:protected="true" table:print="false">
      <table:shapes><draw:rect svg:width="1cm" svg:height="1cm" svg:x="0cm" svg:y="0cm"/></table:shapes>
      <table:table-column/>
      <table:table-row><table:table-cell office:value-type="float" office:value="5"><office:annotation><dc:creator>Ann</dc:creator><text:p>Check this</text:p></office:annotation><text:p>5</text:p></table:table-cell></table:table-row>
      <table:named-expressions><table:named-expression table:name="LocalRate" table:expression="of:=0.05" table:base-cell-address="$Sheet1.$A$1"/><table:named-range table:name="LocalCell" table:cell-range-address="$Sheet1.$A$1" table:base-cell-address="$Sheet1.$A$1"/></table:named-expressions>
      <calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="Sheet1.A1:Sheet1.A1"><calcext:condition calcext:apply-style-name="Bad" calcext:value="&gt;3" calcext:base-cell-address="Sheet1.A1"/></calcext:conditional-format></calcext:conditional-formats>
    </table:table>
    <table:named-expressions><table:named-expression table:name="Rate" table:expression="of:=0.07" table:base-cell-address="$Sheet1.$A$1"/></table:named-expressions>`;

  it('keeps annotations, shapes, conditional formats, table attributes and named expressions', () => {
    const saved = resave(ods(tables));
    const content = readEntry(saved, OdsXmlNames.contentPart);
    expect(content).toContain('<dc:creator>Ann</dc:creator>');
    expect(content).toContain('Check this');
    expect(content).toContain('<table:shapes>');
    expect(content).toContain('calcext:conditional-format ');
    expect(content).toContain('table:protected="true"');
    expect(content).toContain('table:print="false"');
    expect(content).toContain('table:name="Rate" table:expression="of:=0.07"');
    expect(content).toContain('table:name="LocalRate"');
    // The cell's own text is not swallowed by the annotation's paragraph.
    const reopened = readOds(saved);
    expect(reopened.activeSheet.getValue(0, 0)).toEqual(CellValue.number(5));
    expect(reopened.findNamedRange('LocalCell', 'Sheet1')?.scope).toBe('Sheet1');
    // A second save still has them: preservation is not single-use.
    const again = readEntry(resave(saved), OdsXmlNames.contentPart);
    expect(again).toContain('Check this');
    expect(again).toContain('table:name="Rate" table:expression="of:=0.07"');
  });
});

describe('ODS number formats', () => {
  it.each(['hh:mm:ss', 'h:mm AM/PM', '[h]:mm', 'yyyy-mm-dd hh:mm', 'dddd, mmmm d, yyyy', 'mm:ss.00'])(
    'time and date-time code %s round trips',
    (code) => {
      expect(formatAfterRoundTrip(code)).toBe(code);
    },
  );

  it('writes times as a time-style with hour/minute/second elements, never months', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.setValue(0, 0, CellValue.number(0.5));
    workbook.activeSheet.getOrCreate(0, 0).styleId = workbook.styles.intern({
      ...CellStyles.default,
      numberFormat: NumberFormats.of('[h]:mm:ss'),
    });
    const content = contentOf(workbook);
    expect(content).toMatch(/<number:time-style [^>]*number:truncate-on-overflow="false"/);
    expect(content).toContain('<number:hours');
    expect(content).toContain('<number:minutes number:style="long"/>');
    expect(content).toContain('<number:seconds number:style="long"/>');
    expect(content).not.toContain('<number:month');
  });

  it('a currency code keeps its sections, colour and locale', () => {
    const code = '[$$-409]#,##0.00;[RED]-[$$-409]#,##0.00';
    expect(formatAfterRoundTrip(code)).toBe(code);
    const workbook = Workbook.createNew();
    workbook.activeSheet.setValue(0, 0, CellValue.number(-3));
    workbook.activeSheet.getOrCreate(0, 0).styleId = workbook.styles.intern({
      ...CellStyles.default,
      numberFormat: NumberFormats.of(code),
    });
    const content = contentOf(workbook);
    expect(content).toContain('number:language="en" number:country="US"');
    expect(content).toMatch(
      /<style:map style:condition="value\(\)&gt;=0" style:apply-style-name="N\d+P0"\/>/,
    );
    expect(content).toContain('fo:color="#ff0000"');
  });

  it('a three-section code and a euro suffix round trip', () => {
    // The zero section's "-" needs no quotes on the way back: a dash is always literal.
    expect(formatAfterRoundTrip('#,##0.00 [$€-407];-#,##0.00 [$€-407];"-"')).toBe(
      '#,##0.00 [$€-407];-#,##0.00 [$€-407];-',
    );
    expect(formatAfterRoundTrip('$#,##0.00;($#,##0.00)')).toBe('$#,##0.00;($#,##0.00)');
  });

  it('reads a LibreOffice currency style through its style:map', () => {
    const bytes = ods(
      `<table:table table:name="Sheet1"><table:table-row><table:table-cell table:style-name="ce1" office:value-type="currency" office:currency="USD" office:value="-2"><text:p>-$2.00</text:p></table:table-cell></table:table-row></table:table>`,
      {
        automaticStyles: `<number:currency-style style:name="N5P0" style:volatile="true"><number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol><number:number number:decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:currency-style>
          <number:currency-style style:name="N5"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text><number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol><number:number number:decimal-places="2" number:min-integer-digits="1" number:grouping="true"/><style:map style:condition="value()&gt;=0" style:apply-style-name="N5P0"/></number:currency-style>
          <style:style style:name="ce1" style:family="table-cell" style:data-style-name="N5"/>`,
      },
    );
    const workbook = readOds(bytes);
    const code = workbook.styles.get(workbook.activeSheet.find(0, 0)!.styleId).numberFormat.code;
    expect(code).toBe('[$$-409]#,##0.00;[RED]-[$$-409]#,##0.00');
  });

  it('resolves data styles, parent cell styles and font faces from styles.xml', () => {
    const styles = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${NS}><office:font-face-decls><style:font-face style:name="Georgia1" svg:font-family="'Georgia'"/></office:font-face-decls>
<office:styles>
  <number:number-style style:name="N9"><number:number number:decimal-places="3" number:min-integer-digits="1" number:grouping="true"/></number:number-style>
  <style:style style:name="Default" style:family="table-cell"><style:text-properties style:font-name="Georgia1" fo:font-size="12pt"/></style:style>
  <style:style style:name="Heading" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N9"><style:text-properties fo:font-weight="bold"/></style:style>
</office:styles></office:document-styles>`;
    const bytes = ods(
      `<table:table table:name="Sheet1"><table:table-row><table:table-cell table:style-name="ce1" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row></table:table>`,
      {
        styles,
        automaticStyles: `<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Heading"><style:text-properties fo:font-style="italic"/></style:style>`,
      },
    );
    const workbook = readOds(bytes);
    const style = workbook.styles.get(workbook.activeSheet.find(0, 0)!.styleId);
    expect(style.numberFormat.code).toBe('#,##0.000');
    expect(style.font.family).toBe('Georgia');
    expect(style.font.sizePoints).toBe(12);
    expect(style.font.bold).toBe(true);
    expect(style.font.italic).toBe(true);
  });
  it('keeps the spaces of a data style in styles.xml', () => {
    // LibreOffice moves data styles to styles.xml; a whitespace-only <number:text> is a separator.
    const styles = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${NS}><office:styles>
  <number:date-style style:name="N1"><number:day/><number:text> </number:text><number:month number:style="long" number:textual="true"/><number:text> </number:text><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/></number:date-style>
  <style:style style:name="Dated" style:family="table-cell" style:data-style-name="N1"/>
</office:styles></office:document-styles>`;
    const bytes = ods(
      `<table:table table:name="Sheet1"><table:table-row><table:table-cell table:style-name="Dated" office:value-type="date" office:date-value="2024-01-02T03:04:00"><text:p>x</text:p></table:table-cell></table:table-row></table:table>`,
      { styles },
    );
    const workbook = readOds(bytes);
    expect(workbook.styles.get(workbook.activeSheet.find(0, 0)!.styleId).numberFormat.code).toBe(
      'd mmmm hh:mm',
    );
  });
});

describe('ODS addresses and database ranges', () => {
  it('quotes sheet names that need it, doubling apostrophes', () => {
    const workbook = Workbook.createNew();
    workbook.renameSheet(workbook.activeSheet, "It's mine");
    workbook.activeSheet.setValue(0, 0, CellValue.number(1));
    workbook.activeSheet.activeFilter = CellRange.parse('A1:B3');
    workbook.addNamedRange({ name: 'Spot', target: CellRange.parse('A1').withSheet("It's mine") });
    const content = contentOf(workbook);
    expect(content).toContain("$'It''s mine'.$A$1");
    const reopened = roundTrip(workbook);
    expect(reopened.activeSheet.activeFilter?.equals(CellRange.parse('A1:B3'))).toBe(true);
    expect(reopened.findNamedRange('Spot')?.target.sheet).toBe("It's mine");
  });

  it('names each sheet AutoFilter __Anonymous_Sheet_DB__<index> and reads old and new names', () => {
    const workbook = Workbook.createNew();
    workbook.addSheet('Second');
    workbook.worksheets[0]!.activeFilter = CellRange.parse('A1:A4');
    workbook.worksheets[1]!.activeFilter = CellRange.parse('B1:B4');
    const content = contentOf(workbook);
    expect(content).toContain('table:name="__Anonymous_Sheet_DB__0"');
    expect(content).toContain('table:name="__Anonymous_Sheet_DB__1"');
    const reopened = roundTrip(workbook);
    expect(reopened.worksheets[1]!.activeFilter?.equals(CellRange.parse('B1:B4'))).toBe(true);
    expect(reopened.worksheets[0]!.tables).toHaveLength(0);

    for (const name of ['__Anonymous_Sheet__', '__Anonymous_Sheet_DB__7']) {
      const bytes = ods(
        `<table:table table:name="Sheet1"><table:table-row>${float(1)}</table:table-row></table:table>`,
        {
          afterTables: `<table:database-ranges><table:database-range table:name="${name}" table:target-range-address="Sheet1.A1:Sheet1.A3" table:display-filter-buttons="true"/></table:database-ranges>`,
        },
      );
      const sheet = readOds(bytes).activeSheet;
      expect(sheet.activeFilter?.equals(CellRange.parse('A1:A3'))).toBe(true);
      expect(sheet.tables).toHaveLength(0);
    }
  });

  it('a numeric filter condition is typed as a number', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.activeFilter = CellRange.parse('A1:A4');
    sheet.activeFilterSpec = {
      hasHeader: true,
      columns: new Map([[0, { allowedValues: null, condition: '>5' }]]),
    } as never;
    expect(contentOf(workbook)).toMatch(/table:operator="&gt;" table:value="5" table:data-type="number"/);
  });
});

describe('ODS text', () => {
  it('expands text:s, text:tab, text:line-break and spans, keeping whitespace-only runs', () => {
    const bytes = ods(
      `<table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="string"><text:p>a<text:s text:c="3"/>b<text:tab/>c<text:line-break/>d<text:span> <text:s/>e</text:span></text:p><text:p> </text:p></table:table-cell></table:table-row></table:table>`,
    );
    expect(readOds(bytes).activeSheet.getValue(0, 0)).toEqual(CellValue.text('a   b\tc\nd  e\n '));
  });

  it('writes runs of spaces and tabs so they survive', () => {
    const workbook = Workbook.createNew();
    const text = '  two  lead\tand trail  \n second';
    workbook.activeSheet.setValue(0, 0, CellValue.text(text));
    const content = contentOf(workbook);
    expect(content).toContain('<text:s');
    expect(content).toContain('<text:tab/>');
    expect(roundTrip(workbook).activeSheet.getValue(0, 0)).toEqual(CellValue.text(text));
  });
});

describe('ODS validations', () => {
  it('splits an L-shaped target into exact rectangles', () => {
    const cell = '<table:table-cell table:content-validation-name="val1"/>';
    const bytes = ods(
      `<table:table table:name="Sheet1">
        <table:table-row>${cell}${cell}</table:table-row>
        <table:table-row>${cell}<table:table-cell/></table:table-row>
        <table:table-row table:number-rows-repeated="3">${cell}${cell}</table:table-row>
      </table:table>`,
      {
        beforeTables: `<table:content-validations><table:content-validation table:name="val1" table:condition="of:cell-content-is-whole-number() and cell-content()&gt;1" table:base-cell-address="Sheet1.A1"/></table:content-validations>`,
      },
    );
    const targets = readOds(bytes)
      .activeSheet.validations.map((v) => v.target.toString())
      .sort();
    expect(targets).toEqual(['A1:B1', 'A2', 'A3:B5'].map((a) => CellRange.parse(a).toString()).sort());
  });

  it('keeps a condition it cannot model verbatim', () => {
    const condition = 'of:is-true-formula(AND([.A1]&gt;0;ISEVEN([.A1])))';
    const saved = resave(
      ods(
        `<table:table table:name="Sheet1"><table:table-row><table:table-cell table:content-validation-name="v"/></table:table-row></table:table>`,
        {
          beforeTables: `<table:content-validations><table:content-validation table:name="v" table:condition="${condition}" table:base-cell-address="Sheet1.A1"/></table:content-validations>`,
        },
      ),
    );
    expect(readEntry(saved, OdsXmlNames.contentPart)).toContain(`table:condition="${condition}"`);
  });

  it('a whole-column validation stays small on disk', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.setValue(0, 0, CellValue.number(1));
    workbook.activeSheet.addValidation(
      dataValidation('wholeNumber', CellRange.parse('B1:B1048576'), {
        operator: 'between',
        operand1: 1,
        operand2: 9,
      }),
    );
    const content = contentOf(workbook);
    expect(content.length).toBeLessThan(20_000);
    const rules = roundTrip(workbook).activeSheet.validations;
    expect(rules).toHaveLength(1);
    expect(rules[0]!.target.equals(CellRange.parse('B1:B1048576'))).toBe(true);
  });
});

describe('ODS values', () => {
  it('writes fractional seconds in a date-time value', () => {
    const workbook = Workbook.createNew();
    const serial = 45000 + (12 * 3600 + 0.25) / 86400;
    workbook.activeSheet.setValue(0, 0, CellValue.number(serial));
    workbook.activeSheet.getOrCreate(0, 0).styleId = workbook.styles.intern({
      ...CellStyles.default,
      numberFormat: NumberFormats.of('yyyy-mm-dd hh:mm:ss.00'),
    });
    expect(contentOf(workbook)).toMatch(/office:date-value="\d{4}-\d\d-\d\dT12:00:00\.25"/);
    const back = roundTrip(workbook).activeSheet.getValue(0, 0);
    expect(back.kind === 'number' && Math.abs(back.value - serial) < 1e-9).toBe(true);
  });

  it('reads office:boolean-value="1" as TRUE', () => {
    const bytes = ods(
      `<table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="boolean" office:boolean-value="1"><text:p>TRUE</text:p></table:table-cell></table:table-row></table:table>`,
    );
    expect(readOds(bytes).activeSheet.getValue(0, 0)).toEqual(CellValue.boolean(true));
  });

  it('drops XML-invalid control characters from a chart title', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.setValue(0, 0, CellValue.number(1));
    const chart = createChart({ title: 'Q\u0001 sales\u001f', type: 'column' });
    chart.series.push(createSeries({ valuesReference: 'Sheet1.A1:A1' }));
    workbook.addChartSheet(chart);
    const bytes = writeOds(workbook);
    const objectContent = entryNames(bytes).find((n) => /^Object \d+\/content\.xml$/.test(n))!;
    const chartXml = readEntry(bytes, objectContent);
    expect(chartXml.includes('\u0001') || chartXml.includes('\u001f')).toBe(false);
    expect(readOds(bytes).findSheet('Chart1')!.chart!.title).toBe('Q sales');
  });

  it('keeps the manifest entry of a preserved embedded object', () => {
    const manifest = `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="Object 9/" manifest:media-type="application/vnd.oasis.opendocument.formula"/><manifest:file-entry manifest:full-path="Object 9/content.xml" manifest:media-type="text/xml"/></manifest:manifest>`;
    const saved = resave(
      ods(`<table:table table:name="Sheet1"><table:table-row>${float(1)}</table:table-row></table:table>`, {
        manifest,
        extra: { 'Object 9/content.xml': '<math/>' },
      }),
    );
    expect(readEntry(saved, OdsXmlNames.manifestPart)).toContain(
      'manifest:full-path="Object 9/" manifest:media-type="application/vnd.oasis.opendocument.formula"',
    );
  });
});

describe('CSV', () => {
  const invariant = { ...CsvOptions.default, locale: 'en-US' };

  it('exports General numbers with 15 significant digits', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.number(1e-20));
    sheet.setValue(0, 1, CellValue.number(0.1 + 0.2));
    sheet.setValue(0, 2, CellValue.number(1 / 3));
    sheet.setValue(0, 3, CellValue.number(123456789.1234567));
    expect(writeCsv(workbook, sheet, invariant)).toBe('1e-20,0.3,0.333333333333333,123456789.123457\n');
    expect(writeCsv(workbook, sheet, { ...CsvOptions.default, delimiter: ';', locale: 'de-DE' })).toBe(
      '1e-20;0,3;0,333333333333333;123456789,123457\n',
    );
  });

  it('visits only occupied cells of a huge sparse sheet', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.text('a'));
    sheet.setValue(2, 16383, CellValue.text('z'));
    const started = performance.now();
    const csv = writeCsv(workbook, sheet, invariant);
    expect(performance.now() - started).toBeLessThan(2_000);
    const lines = csv.split('\n');
    expect(lines).toHaveLength(4);
    expect(lines[0]).toBe('a' + ','.repeat(16383));
    expect(lines[1]).toBe(','.repeat(16383));
    expect(lines[2]).toBe(','.repeat(16383) + 'z');
  });
});
