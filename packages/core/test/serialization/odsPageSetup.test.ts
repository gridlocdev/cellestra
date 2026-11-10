import { describe, expect, it } from 'vitest';
import { SetPageSetupCommand } from '../../src/editing/pageSetupCommand.ts';
import { PageSetups, type PageSetup } from '../../src/model/pageSetup.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { OdsWriter } from '../../src/serialization/ods/odsWriter.ts';
import { OdsPackage, OdsXmlNames, readOdsWithPackage } from '../../src/serialization/index.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { readEntry } from './helpers.ts';

const landscapeLetter: PageSetup = {
  ...PageSetups.orient({ ...PageSetups.default, paperWidth: 215.9, paperHeight: 279.4 }, 'landscape'),
  margins: { top: 10, bottom: 15, left: 12, right: 12 },
  scale: { kind: 'fit', width: 1, height: 0 },
  gridlines: true,
  headers: true,
  centerHorizontally: true,
};

function workbook(): Workbook {
  const wb = Workbook.createNew();
  wb.worksheets[0]!.setValue(0, 0, CellValue.number(1));
  return wb;
}

const NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"';

const count = (text: string, needle: string) => text.split(needle).length - 1;

describe('page setup', () => {
  it('is saved as a page layout and read back', () => {
    const wb = workbook();
    wb.worksheets[0]!.pageSetup = landscapeLetter;
    const back = readOdsWithPackage(OdsWriter.write(wb)).workbook;
    expect(PageSetups.equal(back.worksheets[0]!.pageSetup, landscapeLetter)).toBe(true);
  });

  it('reads each sheet its own page, and the default page for sheets without one', () => {
    const wb = workbook();
    wb.addSheet('Two');
    wb.worksheets[0]!.pageSetup = landscapeLetter;
    wb.worksheets[1]!.pageSetup = { ...PageSetups.default, scale: { kind: 'percent', percent: 75 } };
    const back = readOdsWithPackage(OdsWriter.write(wb)).workbook;
    expect(back.worksheets[0]!.pageSetup?.orientation).toBe('landscape');
    expect(back.worksheets[1]!.pageSetup?.scale).toEqual({ kind: 'percent', percent: 75 });
  });

  it('saving again replaces the page it wrote instead of adding another', () => {
    const wb = workbook();
    wb.worksheets[0]!.pageSetup = landscapeLetter;
    const first = readOdsWithPackage(OdsWriter.write(wb));
    first.workbook.worksheets[0]!.pageSetup = { ...landscapeLetter, gridlines: false };
    const bytes = OdsWriter.write(first.workbook, first.package);
    const styles = readEntry(bytes, OdsXmlNames.stylesPart);
    const content = readEntry(bytes, OdsXmlNames.contentPart);
    expect(count(styles, '<style:master-page style:name="CellestraPage')).toBe(1);
    expect(count(styles, 'style:name="CellestraPL')).toBe(1);
    expect(count(content, 'style:name="CellestraTa')).toBe(1);
    const again = readOdsWithPackage(bytes).workbook;
    expect(again.worksheets[0]!.pageSetup?.gridlines).toBe(false);
  });

  it("leaves a file's own page alone when it did not change", () => {
    const pkg = new OdsPackage();
    pkg.setText(OdsXmlNames.mimetypePart, OdsXmlNames.mimeType);
    pkg.setText(
      OdsXmlNames.contentPart,
      `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3">` +
        `<office:automatic-styles><style:style style:name="ta1" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true"/></style:style></office:automatic-styles>` +
        `<office:body><office:spreadsheet><table:table table:name="Sheet1" table:style-name="ta1"><table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`,
    );
    const styles =
      `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS} office:version="1.3">` +
      `<office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="8.5in" fo:page-height="11in" style:print-orientation="portrait" fo:margin="0.5in" style:scale-to="80%" style:print="grid"/></style:page-layout></office:automatic-styles>` +
      `<office:master-styles><style:master-page style:name="Default" style:page-layout-name="pm1"><style:header><text:p>Head</text:p></style:header></style:master-page></office:master-styles></office:document-styles>`;
    pkg.setText(OdsXmlNames.stylesPart, styles);
    const read = readOdsWithPackage(pkg.write());
    const setup = read.workbook.worksheets[0]!.pageSetup!;
    expect(setup.paperWidth).toBeCloseTo(215.9);
    expect(setup.margins.left).toBeCloseTo(12.7);
    expect(setup.scale).toEqual({ kind: 'percent', percent: 80 });
    expect(setup.gridlines).toBe(true);
    const unchanged = OdsWriter.write(read.workbook, read.package);
    expect(readEntry(unchanged, OdsXmlNames.stylesPart)).toBe(styles);

    // Changed, the new page copies the old one's header.
    const again = readOdsWithPackage(unchanged);
    again.workbook.worksheets[0]!.pageSetup = { ...setup, gridlines: false };
    const changed = readEntry(OdsWriter.write(again.workbook, again.package), OdsXmlNames.stylesPart);
    expect(changed).toContain(
      '<style:master-page style:name="CellestraPage0" style:page-layout-name="CellestraPL0"><style:header>',
    );
    expect(changed).toContain('style:name="pm1"');
  });

  it('is undoable', () => {
    const wb = workbook();
    const command = new SetPageSetupCommand('Sheet1', landscapeLetter);
    command.apply(wb);
    expect(wb.worksheets[0]!.pageSetup).toBe(landscapeLetter);
    command.undo(wb);
    expect(wb.worksheets[0]!.pageSetup).toBeNull();
  });

  it('turns paper with the orientation and recognises named sizes', () => {
    const turned = PageSetups.orient(PageSetups.default, 'landscape');
    expect([turned.paperWidth, turned.paperHeight]).toEqual([297, 210]);
    expect(PageSetups.paperOf(turned)?.id).toBe('a4');
    expect(PageSetups.paperOf({ ...turned, paperWidth: 100 })).toBeNull();
  });
});
