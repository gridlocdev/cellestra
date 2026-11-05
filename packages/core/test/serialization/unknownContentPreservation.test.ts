import { describe, expect, it } from 'vitest';
import { OdsPackage, OdsXmlNames, readOdsWithPackage, writeOds } from '../../src/serialization/index.ts';
import { Office, Table, Text } from '../../src/serialization/ods/odsXmlNames.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { serializeXml, XmlDocument, XmlElement } from '../../src/xml/index.ts';

const EXTRA_PART_BYTES = new Uint8Array([1, 2, 3, 4, 5]);

/** Hand-built .ods with an unmodeled spreadsheet child and an extra binary part. */
function buildPackageWithUnknownContent(): Uint8Array {
  const content = new XmlElement(Office('document-content'))
    .declare(Office.prefix, Office.uri)
    .declare(Table.prefix, Table.uri)
    .declare(Text.prefix, Text.uri)
    .add(
      new XmlElement(
        Office('body'),
        new XmlElement(
          Office('spreadsheet'),
          new XmlElement(Table('table'))
            .setAttr(Table('name'), 'Sheet1')
            .add(
              new XmlElement(
                Table('table-row'),
                new XmlElement(Table('table-cell'), new XmlElement(Text('p'), '5'))
                  .setAttr(Office('value-type'), 'float')
                  .setAttr(Office('value'), '5'),
              ),
            ),
          // An element Cellestra does not model (a pivot table) — must survive.
          new XmlElement(
            Table('data-pilot-tables'),
            new XmlElement(Table('data-pilot-table')).setAttr(Table('name'), 'PT'),
          ),
        ),
      ),
    );

  const pkg = new OdsPackage();
  pkg.setText(OdsXmlNames.mimetypePart, OdsXmlNames.mimeType);
  pkg.setText(OdsXmlNames.contentPart, serializeXml(new XmlDocument(content)));
  pkg.setBytes('Pictures/logo.bin', EXTRA_PART_BYTES);
  return pkg.write();
}

describe('UnknownContentPreservation', () => {
  it('unknown spreadsheet child and extra part survive a round trip', () => {
    // Open the crafted file.
    const { workbook, package: pkg } = readOdsWithPackage(buildPackageWithUnknownContent());

    expect(workbook.findSheet('Sheet1')!.getValue(0, 0)).toEqual(CellValue.number(5));
    expect(pkg.foreignSpreadsheetChildren.some((f) => f.xml.includes('data-pilot-tables'))).toBe(true);
    expect(pkg.contains('Pictures/logo.bin')).toBe(true);

    // Save it back through the model, then re-open and confirm nothing was dropped.
    const { workbook: workbook2, package: pkg2 } = readOdsWithPackage(writeOds(workbook, pkg));

    expect(workbook2.findSheet('Sheet1')!.getValue(0, 0)).toEqual(CellValue.number(5));
    expect(pkg2.foreignSpreadsheetChildren.some((f) => f.xml.includes('data-pilot-tables'))).toBe(true);
    expect(pkg2.contains('Pictures/logo.bin')).toBe(true);
    expect(pkg2.getBytes('Pictures/logo.bin')).toEqual(EXTRA_PART_BYTES);
  });

  it('a foreign child keeps its position and its exact markup', () => {
    // Beyond the C# test: the child is re-emitted before the tables when it came before them, and
    // byte-for-byte as read (attributes, order, whitespace inside it).
    const foreign =
      '<table:calculation-settings xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" table:automatic-find-labels="false"> <table:iteration table:status="enable"/></table:calculation-settings>';
    const content = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet>${foreign.replace(' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"', '')}<table:table table:name="Sheet1"><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const pkg = new OdsPackage();
    pkg.setText(OdsXmlNames.mimetypePart, OdsXmlNames.mimeType);
    pkg.setText(OdsXmlNames.contentPart, content);

    const first = readOdsWithPackage(pkg.write());
    expect(first.package.foreignSpreadsheetChildren).toEqual([{ beforeTables: true, xml: foreign }]);

    const second = readOdsWithPackage(writeOds(first.workbook, first.package));
    expect(second.package.foreignSpreadsheetChildren).toEqual([{ beforeTables: true, xml: foreign }]);
  });
});
