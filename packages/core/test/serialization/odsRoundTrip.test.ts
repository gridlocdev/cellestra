import { describe, expect, it } from 'vitest';
import { DateConversion } from '../../src/formulas/evaluation/dateConversion.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { MergedRegion } from '../../src/model/mergedRegion.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { OdsSupportXml, writeOds } from '../../src/serialization/index.ts';
import {
  BorderEdges,
  CellStyles,
  Fills,
  Fonts,
  NumberFormats,
  type Border,
} from '../../src/styling/cellStyle.ts';
import { Argb } from '../../src/styling/color.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { readEntry, roundTrip } from './helpers.ts';

function buildSample(): Workbook {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;

  sheet.setValue(0, 0, CellValue.number(42));
  sheet.setValue(0, 1, CellValue.text('hi'));
  sheet.setValue(0, 2, CellValue.boolean(true));
  sheet.setFormula(1, 0, 'A1*2', CellValue.number(84));

  const styleId = workbook.styles.intern({
    ...CellStyles.default,
    font: { ...Fonts.default, bold: true, strikethrough: true },
    fill: Fills.solid(Argb.fromRgb(255, 0, 0)),
    numberFormat: NumberFormats.of('0.00'),
  });
  sheet.setValue(0, 3, CellValue.number(3.14159));
  sheet.getOrCreate(0, 3).styleId = styleId;

  sheet.getOrCreateColumn(0).width = 100;
  sheet.getOrCreateRow(0).height = 30;
  sheet.getOrCreateColumn(2).hidden = true;
  sheet.getOrCreateRow(3).hidden = true;

  sheet.addMerge(new MergedRegion(CellRange.parse('A5:B6')));
  workbook.addNamedRange({ name: 'Rate', target: CellRange.parse('Sheet1.A1') });

  const sheet2 = workbook.addSheet();
  sheet2.setValue(0, 0, CellValue.number(7));

  workbook.settings.activeSheetIndex = 1;
  workbook.settings.calculationMode = 'manual';
  workbook.settings.locale = 'de-DE';
  return workbook;
}

function numberAt(workbook: Workbook, sheetName: string, row: number, column: number): number {
  return CellValue.asNumber(workbook.findSheet(sheetName)!.getValue(row, column));
}

describe('OdsRoundTrip', () => {
  it('values and formula survive', () => {
    const sheet = roundTrip(buildSample()).findSheet('Sheet1')!;

    expect(sheet.getValue(0, 0)).toEqual(CellValue.number(42));
    expect(sheet.getValue(0, 1)).toEqual(CellValue.text('hi'));
    expect(sheet.getValue(0, 2)).toEqual(CellValue.boolean(true));

    const formulaCell = sheet.find(1, 0)!;
    expect(formulaCell.formula).toBe('A1*2');
    expect(formulaCell.value).toEqual(CellValue.number(84)); // cached value loaded for instant display
  });

  it('date time of day survives round trip', () => {
    // A datetime cell carries a fractional serial; saving must not floor it to midnight.
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    const serial = DateConversion.toSerial({
      year: 2026,
      month: 6,
      day: 13,
      hour: 15,
      minute: 30,
      second: 0,
      millisecond: 0,
    });

    const dateTimeStyle = workbook.styles.intern({
      ...CellStyles.default,
      numberFormat: NumberFormats.of('YYYY-MM-DD HH:MM:SS'),
    });
    sheet.setValue(0, 0, CellValue.number(serial));
    sheet.getOrCreate(0, 0).styleId = dateTimeStyle;

    // A pure time-of-day cell (noon = 0.5) and an elapsed time past 24h must both survive.
    const timeStyle = workbook.styles.intern({
      ...CellStyles.default,
      numberFormat: NumberFormats.of('HH:MM:SS'),
    });
    sheet.setValue(1, 0, CellValue.number(0.5));
    sheet.getOrCreate(1, 0).styleId = timeStyle;
    sheet.setValue(2, 0, CellValue.number(1.25));
    sheet.getOrCreate(2, 0).styleId = timeStyle;

    const result = roundTrip(workbook);

    expect(numberAt(result, 'Sheet1', 0, 0)).toBeCloseTo(serial, 6);
    expect(numberAt(result, 'Sheet1', 1, 0)).toBeCloseTo(0.5, 6);
    expect(numberAt(result, 'Sheet1', 2, 0)).toBeCloseTo(1.25, 6);
  });

  it('non-finite numbers serialize as valid ODF and survive', () => {
    // office:value is an xsd:double whose lexical space is INF/-INF/NaN.
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.number(Infinity));
    sheet.setValue(1, 0, CellValue.number(-Infinity));
    sheet.setValue(2, 0, CellValue.number(NaN));

    const bytes = writeOds(workbook);
    const content = readEntry(bytes, 'content.xml');
    expect(content).toContain('office:value="INF"');
    expect(content).toContain('office:value="-INF"');
    expect(content).toContain('office:value="NaN"');
    expect(content).not.toContain('office:value="Infinity"');
    expect(content).not.toContain('office:value="-Infinity"');

    const reopened = roundTrip(workbook);
    expect(numberAt(reopened, 'Sheet1', 0, 0)).toBe(Infinity);
    expect(numberAt(reopened, 'Sheet1', 1, 0)).toBe(-Infinity);
    expect(numberAt(reopened, 'Sheet1', 2, 0)).toBeNaN();
  });

  it('styles survive', () => {
    const result = roundTrip(buildSample());
    const sheet = result.findSheet('Sheet1')!;
    const style = result.styles.get(sheet.find(0, 3)!.styleId);

    expect(style.font.bold).toBe(true);
    expect(style.font.strikethrough).toBe(true);
    expect(Fills.hasBackground(style.fill)).toBe(true);
    expect(style.fill.background).toBe(Argb.fromRgb(255, 0, 0));
    expect(style.numberFormat.code).toBe('0.00');
  });

  it('borders survive', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    const border: Border = {
      left: BorderEdges.of('dashed', Argb.fromRgb(0, 0, 255)),
      right: BorderEdges.of('thick'),
      top: BorderEdges.of('thin'),
      bottom: BorderEdges.of('double', Argb.fromRgb(0, 128, 0)),
    };
    const styleId = workbook.styles.intern({ ...CellStyles.default, border });
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.getOrCreate(0, 0).styleId = styleId;

    const result = roundTrip(workbook);
    const b = result.styles.get(result.findSheet('Sheet1')!.find(0, 0)!.styleId).border;

    expect(b.left.style).toBe('dashed');
    expect(b.left.color).toBe(Argb.fromRgb(0, 0, 255));
    expect(b.right.style).toBe('thick');
    expect(b.top.style).toBe('thin');
    expect(b.bottom.style).toBe('double');
    expect(b.bottom.color).toBe(Argb.fromRgb(0, 128, 0));
  });

  it('layout survives', () => {
    const sheet = roundTrip(buildSample()).findSheet('Sheet1')!;

    expect(Math.abs(sheet.getColumnWidth(0) - 100)).toBeLessThanOrEqual(0.5);
    expect(Math.abs(sheet.getRowHeight(0) - 30)).toBeLessThanOrEqual(0.5);
    expect(sheet.isColumnHidden(2)).toBe(true);
    expect(sheet.isRowHidden(3)).toBe(true);
  });

  it('merges survive', () => {
    const sheet = roundTrip(buildSample()).findSheet('Sheet1')!;
    const merge = sheet.findMerge(4, 0);

    expect(merge).toBeDefined();
    expect(merge!.range.equals(CellRange.parse('A5:B6'))).toBe(true);
  });

  it('sheets, named ranges, and settings survive', () => {
    const result = roundTrip(buildSample());

    expect(result.sheetCount).toBe(2);
    expect(numberAt(result, 'Sheet2', 0, 0)).toBe(7);
    expect(result.findNamedRange('Rate')).toBeDefined();
    expect(result.settings.activeSheetIndex).toBe(1);
    expect(result.settings.calculationMode).toBe('manual');
    expect(result.settings.locale).toBe('de-DE'); // the document locale round-trips (settings.xml)
  });

  it('a hand-broken string survives as one paragraph per line', () => {
    // ODF has no newline inside a paragraph, so a multi-line cell is written one text:p per line.
    const workbook = Workbook.createNew();
    workbook.activeSheet.setValue(0, 0, CellValue.text('first\nsecond'));

    const xml = readEntry(writeOds(workbook), 'content.xml');

    expect(xml).not.toContain('first\nsecond');
    expect(roundTrip(workbook).activeSheet.getValue(0, 0)).toEqual(CellValue.text('first\nsecond'));
  });

  it('a hyperlink is written unindented inside its paragraph, and formulas resolve the of: prefix', () => {
    // Whitespace inside text:p is cell text to a reader, so the anchor must sit flush.
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.text('link'));
    sheet.getOrCreate(0, 0).hyperlink = { target: 'https://example.com' };

    const xml = readEntry(writeOds(workbook), 'content.xml');

    expect(xml).toContain('<text:p><text:a xlink:href="https://example.com">link</text:a></text:p>');
    expect(xml).toContain('xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"');
    const cell = roundTrip(workbook).activeSheet.getOrCreate(0, 0);
    expect(cell.value).toEqual(CellValue.text('link'));
    expect(cell.hyperlink?.target).toBe('https://example.com');
  });

  it('frozen panes survive', () => {
    const workbook = buildSample();
    workbook.findSheet('Sheet1')!.frozenRows = 1;
    workbook.findSheet('Sheet1')!.frozenColumns = 2;

    const result = roundTrip(workbook);

    expect(result.findSheet('Sheet1')!.frozenRows).toBe(1);
    expect(result.findSheet('Sheet1')!.frozenColumns).toBe(2);
    expect(result.findSheet('Sheet2')!.frozenRows).toBe(0);
  });

  it('a movable split is not read as a freeze', () => {
    const workbook = Workbook.createNew();
    OdsSupportXml.applySettings(
      `<?xml version="1.0" encoding="UTF-8"?>
<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                          xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">
  <office:settings>
    <config:config-item-set config:name="ooo:view-settings">
      <config:config-item-map-indexed config:name="Views">
        <config:config-item-map-entry>
          <config:config-item-map-named config:name="Tables">
            <config:config-item-map-entry config:name="Sheet1">
              <config:config-item config:name="VerticalSplitMode" config:type="short">1</config:config-item>
              <config:config-item config:name="VerticalSplitPosition" config:type="int">3</config:config-item>
            </config:config-item-map-entry>
          </config:config-item-map-named>
        </config:config-item-map-entry>
      </config:config-item-map-indexed>
    </config:config-item-set>
  </office:settings>
</office:document-settings>`,
      workbook,
    );

    expect(workbook.activeSheet.frozenRows).toBe(0);
  });

  it('mimetype is first and stored uncompressed', () => {
    const bytes = writeOds(buildSample());
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

    // The first local file header: signature PK\3\4, compression method at offset 8, name at 30.
    expect(view.getUint32(0, true)).toBe(0x04034b50);
    expect(view.getUint16(8, true)).toBe(0); // stored, not deflated
    const nameLength = view.getUint16(26, true);
    const name = String.fromCharCode(...bytes.subarray(30, 30 + nameLength));
    expect(name).toBe('mimetype');
    const compressedSize = view.getUint32(18, true);
    const uncompressedSize = view.getUint32(22, true);
    expect(compressedSize).toBe(uncompressedSize);
    const extraLength = view.getUint16(28, true);
    const dataStart = 30 + nameLength + extraLength;
    const payload = String.fromCharCode(...bytes.subarray(dataStart, dataStart + uncompressedSize));
    expect(payload).toBe('application/vnd.oasis.opendocument.spreadsheet');
  });
});
