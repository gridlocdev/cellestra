import { describe, expect, it } from 'vitest';
import { Workbook } from '../../src/model/workbook.ts';
import { OdsPackage, OdsXmlNames, readOdsWithPackage, writeOds } from '../../src/serialization/index.ts';
import { CellStyles, Protections, styleKey, type CellStyle } from '../../src/styling/cellStyle.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { parseXml, type XmlElement } from '../../src/xml/index.ts';
import { Fo, Office, Style, Table } from '../../src/serialization/ods/odsXmlNames.ts';
import { readEntry, roundTrip } from './helpers.ts';

/**
 * Cell protection, text rotation and legacy conditional formats (`style:map`) are properties of a
 * cell style the grid does not draw, but a save must keep them, per cell.
 */

const NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"';

/** A file whose cells A1… use `cellStyles` (ce1, ce2, …) in order, one cell per row. */
function buildOds(cellStyles: readonly string[], extraAutomatic = ''): Uint8Array {
  const automatic = cellStyles
    .map(
      (body, i) =>
        `<style:style style:name="ce${i + 1}" style:family="table-cell" style:parent-style-name="Default">${body}</style:style>`,
    )
    .join('');
  const rows = cellStyles
    .map(
      (_, i) =>
        `<table:table-row><table:table-cell table:style-name="ce${i + 1}" office:value-type="float" office:value="${i + 1}"><text:p>${i + 1}</text:p></table:table-cell></table:table-row>`,
    )
    .join('');
  const content =
    `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3">` +
    `<office:automatic-styles>${automatic}${extraAutomatic}</office:automatic-styles>` +
    `<office:body><office:spreadsheet><table:table table:name="Sheet1">${rows}</table:table></office:spreadsheet></office:body>` +
    `</office:document-content>`;
  const styles =
    `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS} office:version="1.3"><office:styles>` +
    `<style:style style:name="Default" style:family="table-cell"/>` +
    `<style:style style:name="Good" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#ccffcc"/></style:style>` +
    `</office:styles></office:document-styles>`;
  const pkg = new OdsPackage();
  pkg.setText(OdsXmlNames.mimetypePart, OdsXmlNames.mimeType);
  pkg.setText(OdsXmlNames.contentPart, content);
  pkg.setText(OdsXmlNames.stylesPart, styles);
  return pkg.write();
}

/** The model's reading of a cell's style, without the file style it was read from. */
function styleOf(workbook: Workbook, row: number): CellStyle {
  const { source: _source, ...style } = workbook.styles.get(
    workbook.findSheet('Sheet1')!.find(row, 0)!.styleId,
  );
  return style;
}

/** content.xml's automatic cell styles by name, and the style name each cell of column A uses. */
function writtenStyles(bytes: Uint8Array): { byName: Map<string, XmlElement>; cellStyleNames: string[] } {
  const root = parseXml(readEntry(bytes, OdsXmlNames.contentPart)).root;
  const byName = new Map<string, XmlElement>();
  for (const s of root.element(Office('automatic-styles'))!.elements(Style('style'))) {
    if (s.attr(Style('family')) === 'table-cell') byName.set(s.attr(Style('name'))!, s);
  }
  // Conditional formats apply named styles, which a save writes into styles.xml.
  const styles = parseXml(readEntry(bytes, OdsXmlNames.stylesPart)).root.element(Office('styles'));
  for (const s of styles?.elements(Style('style')) ?? []) {
    if (s.attr(Style('family')) === 'table-cell') byName.set(s.attr(Style('name'))!, s);
  }
  const cellStyleNames = root
    .descendants(Table('table-cell'))
    .map((c) => c.attr(Table('style-name')))
    .filter((n): n is string => n !== null);
  return { byName, cellStyleNames };
}

function cellProps(bytes: Uint8Array, row: number): XmlElement | null {
  const { byName, cellStyleNames } = writtenStyles(bytes);
  return byName.get(cellStyleNames[row]!)!.element(Style('table-cell-properties'));
}

/** open → save (keeping the package) → bytes. */
function resave(bytes: Uint8Array): Uint8Array {
  const { workbook, package: pkg } = readOdsWithPackage(bytes);
  return writeOds(workbook, pkg);
}

describe('cell style protection, rotation and maps', () => {
  it('reads protection, keeps cells that differ only in it distinct, and writes it back', () => {
    const bytes = buildOds([
      '<style:table-cell-properties style:cell-protect="none" style:print-content="true"/>',
      '<style:table-cell-properties style:cell-protect="hidden-and-protected"/>',
      '<style:table-cell-properties style:cell-protect="protected formula-hidden" style:print-content="false"/>',
      '<style:table-cell-properties style:cell-protect="protected"/>',
    ]);
    const { workbook } = readOdsWithPackage(bytes);

    expect(styleOf(workbook, 0).protection).toEqual({ ...Protections.default, locked: false });
    expect(styleOf(workbook, 1).protection).toEqual({
      locked: true,
      hideFormula: true,
      hideAll: true,
      printContent: true,
    });
    expect(styleOf(workbook, 2).protection).toEqual({
      locked: true,
      hideFormula: true,
      hideAll: false,
      printContent: false,
    });
    // Calc's default protection is the default style.
    expect(styleOf(workbook, 3)).toEqual(CellStyles.default);
    expect(new Set([0, 1, 2, 3].map((r) => workbook.findSheet('Sheet1')!.find(r, 0)!.styleId)).size).toBe(4);

    const saved = resave(bytes);
    expect(cellProps(saved, 0)!.attr(Style('cell-protect'))).toBe('none');
    // The style is written back as the file had it.
    expect(cellProps(saved, 0)!.attr(Style('print-content'))).toBe('true');
    expect(cellProps(saved, 1)!.attr(Style('cell-protect'))).toBe('hidden-and-protected');
    expect(cellProps(saved, 2)!.attr(Style('cell-protect'))).toBe('protected formula-hidden');
    expect(cellProps(saved, 2)!.attr(Style('print-content'))).toBe('false');

    const reopened = readOdsWithPackage(saved).workbook;
    for (const row of [0, 1, 2, 3]) expect(styleOf(reopened, row)).toEqual(styleOf(workbook, row));
  });

  it('reads rotation and its anchor, keeps them distinct, and writes them back', () => {
    const bytes = buildOds([
      '<style:table-cell-properties style:rotation-angle="90" style:rotation-align="none"/>',
      '<style:table-cell-properties style:rotation-angle="45" style:rotation-align="bottom"/>',
      '<style:table-cell-properties style:rotation-angle="45deg" style:rotation-align="top"/>',
      '<style:table-cell-properties style:rotation-angle="-90"/>',
      '<style:table-cell-properties style:rotation-angle="0" style:rotation-align="bottom"/>',
    ]);
    const { workbook } = readOdsWithPackage(bytes);

    expect(styleOf(workbook, 0).alignment).toMatchObject({ rotation: 90, rotationAlign: 'none' });
    expect(styleOf(workbook, 1).alignment).toMatchObject({ rotation: 45, rotationAlign: 'bottom' });
    expect(styleOf(workbook, 2).alignment).toMatchObject({ rotation: 45, rotationAlign: 'top' });
    expect(styleOf(workbook, 3).alignment.rotation).toBe(270);
    // No rotation is the default style, whatever its anchor says.
    expect(styleOf(workbook, 4)).toEqual(CellStyles.default);

    const saved = resave(bytes);
    expect(cellProps(saved, 0)!.attr(Style('rotation-angle'))).toBe('90');
    expect(cellProps(saved, 0)!.attr(Style('rotation-align'))).toBe('none');
    expect(cellProps(saved, 2)!.attr(Style('rotation-align'))).toBe('top');
    // Written back as the file spelt it.
    expect(cellProps(saved, 3)!.attr(Style('rotation-angle'))).toBe('-90');
    expect(cellProps(saved, 3)!.attr(Style('rotation-align'))).toBeNull();

    const reopened = readOdsWithPackage(saved).workbook;
    for (const row of [0, 1, 2, 3, 4]) expect(styleOf(reopened, row)).toEqual(styleOf(workbook, row));
  });

  it('keeps cells that differ only in a style:map distinct, and the common style it names resolves', () => {
    const bytes = buildOds([
      '<style:map style:condition="cell-content()&gt;5" style:apply-style-name="Good" style:base-cell-address="Sheet1.A1"/>',
      '<style:map style:condition="cell-content()&lt;2" style:apply-style-name="Good" style:base-cell-address="Sheet1.A2"/>',
      '',
    ]);
    const { workbook } = readOdsWithPackage(bytes);
    const ids = [0, 1, 2].map((r) => workbook.findSheet('Sheet1')!.find(r, 0)!.styleId);
    expect(new Set(ids).size).toBe(3);
    expect(styleOf(workbook, 0).conditionalMaps).toEqual([
      {
        condition: 'cell-content()>5',
        applyStyleName: 'Good',
        applyStyle: expect.objectContaining({ fill: { background: expect.any(Number) } }),
        baseCellAddress: 'Sheet1.A1',
      },
    ]);
    expect(styleOf(workbook, 2)).toEqual(CellStyles.default);

    const saved = resave(bytes);
    const { byName, cellStyleNames } = writtenStyles(saved);
    const map = byName.get(cellStyleNames[0]!)!.element(Style('map'))!;
    expect(map.attr(Style('condition'))).toBe('cell-content()>5');
    expect(map.attr(Style('base-cell-address'))).toBe('Sheet1.A1');
    // Good lives in the preserved styles.xml, so the name still resolves.
    expect(map.attr(Style('apply-style-name'))).toBe('Good');
    expect(readEntry(saved, OdsXmlNames.stylesPart)).toContain('style:name="Good"');
    expect(byName.get(cellStyleNames[1]!)!.element(Style('map'))!.attr(Style('condition'))).toBe(
      'cell-content()<2',
    );

    const reopened = readOdsWithPackage(saved).workbook;
    for (const row of [0, 1, 2])
      expect(styleKey(styleOf(reopened, row))).toBe(styleKey(styleOf(workbook, row)));
  });

  it('a style:map naming an automatic style points at that style under its new name', () => {
    const bytes = buildOds(
      ['<style:map style:condition="cell-content()=1" style:apply-style-name="ceRed"/>'],
      '<style:style style:name="ceRed" style:family="table-cell"><style:text-properties fo:color="#ff0000"/></style:style>',
    );
    const saved = resave(bytes);
    const { byName, cellStyleNames } = writtenStyles(saved);
    const target = byName.get(cellStyleNames[0]!)!.element(Style('map'))!.attr(Style('apply-style-name'))!;
    // ceRed is not written under its old name; the map names the style it was saved as.
    expect(byName.has('ceRed')).toBe(false);
    expect(byName.get(target)!.element(Style('text-properties'))!.attr(Fo('color'))?.toLowerCase()).toBe(
      '#ff0000',
    );

    const before = styleOf(readOdsWithPackage(bytes).workbook, 0);
    expect(before.conditionalMaps![0]!.applyStyleName).toBeNull();
    const reopened = readOdsWithPackage(saved).workbook;
    expect(styleKey(styleOf(reopened, 0))).toBe(styleKey(before));
  });

  it('a common style a map names is written as a named style when styles.xml is not kept', () => {
    const { workbook } = readOdsWithPackage(
      buildOds(['<style:map style:condition="cell-content()&gt;5" style:apply-style-name="Good"/>']),
    );
    const saved = writeOds(workbook); // no source package: styles.xml is generated afresh
    const { byName, cellStyleNames } = writtenStyles(saved);
    const target = byName.get(cellStyleNames[0]!)!.element(Style('map'))!.attr(Style('apply-style-name'))!;
    expect(byName.has(target)).toBe(true);
  });

  it('saving again rewrites the named styles maps apply instead of adding more', () => {
    const bytes = buildOds(
      ['<style:map style:condition="cell-content()=1" style:apply-style-name="ceRed"/>'],
      '<style:style style:name="ceRed" style:family="table-cell"><style:text-properties fo:color="#ff0000"/></style:style>',
    );
    const once = resave(bytes);
    const twice = resave(once);
    const count = (b: Uint8Array) =>
      (readEntry(b, OdsXmlNames.stylesPart).match(/CellestraCond/g) ?? []).length;
    expect(count(once)).toBe(1);
    expect(count(twice)).toBe(1);
    const { byName, cellStyleNames } = writtenStyles(twice);
    const target = byName.get(cellStyleNames[0]!)!.element(Style('map'))!.attr(Style('apply-style-name'))!;
    expect(byName.has(target)).toBe(true);
  });

  it('a default style writes no protection, rotation or map', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.getOrCreate(0, 0).styleId = workbook.styles.intern({
      ...CellStyles.default,
      font: { ...CellStyles.default.font, bold: true },
      protection: Protections.default,
    });
    const content = readEntry(writeOds(workbook), OdsXmlNames.contentPart);
    expect(content).not.toMatch(/cell-protect|print-content|rotation-|style:map/);
  });

  it('a style built in the model keeps protection and rotation through a save', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.number(1));
    const style: CellStyle = {
      ...CellStyles.default,
      alignment: { ...CellStyles.default.alignment, rotation: 30 },
      protection: { ...Protections.default, locked: false, printContent: false },
    };
    sheet.getOrCreate(0, 0).styleId = workbook.styles.intern(style);
    expect(styleOf(roundTrip(workbook), 0)).toEqual(style);
  });
});
