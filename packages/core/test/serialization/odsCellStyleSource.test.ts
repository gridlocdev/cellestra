import { describe, expect, it } from 'vitest';
import { EditCommands } from '../../src/editing/editCommands.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { OdsXmlNames, readOdsWithPackage, writeOds } from '../../src/serialization/index.ts';
import { Style, Table } from '../../src/serialization/ods/odsXmlNames.ts';
import { BorderEdges } from '../../src/styling/cellStyle.ts';
import { parseXml, type XmlElement } from '../../src/xml/index.ts';
import { odsWith, readEntry } from './helpers.ts';

/**
 * A cell style carries much the model has no field for. A save writes the style the file had, and
 * an edit changes only what it edits.
 */

const RICH_STYLE = `<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Heading">
  <style:table-cell-properties style:diagonal-bl-tr="0.74pt solid #000000" style:shrink-to-fit="true" fo:padding="0.2cm" fo:border-left="2.5pt solid #ff0000" style:text-align-source="fix" style:repeat-content="false"/>
  <style:paragraph-properties fo:text-align="start" fo:margin-left="0.5cm"/>
  <style:text-properties style:text-underline-style="solid" style:text-underline-type="double" style:text-underline-width="auto" style:text-underline-color="font-color" fo:font-variant="small-caps" fo:font-style="italic" style:font-name="Liberation Serif"/>
</style:style>
<style:style style:name="ce2" style:family="table-cell"><style:table-cell-properties style:diagonal-tl-br="0.74pt solid #000000"/></style:style>
<style:style style:name="ce3" style:family="table-cell"><style:table-cell-properties style:diagonal-bl-tr="0.74pt solid #000000"/></style:style>`;

const SHEET = `<table:table table:name="Sheet1">
  <table:table-row><table:table-cell table:style-name="ce1" office:value-type="string"><text:p>a</text:p></table:table-cell></table:table-row>
  <table:table-row><table:table-cell table:style-name="ce2" office:value-type="string"><text:p>b</text:p></table:table-cell></table:table-row>
  <table:table-row><table:table-cell table:style-name="ce3" office:value-type="string"><text:p>c</text:p></table:table-cell></table:table-row>
</table:table>`;

const FONT_FACES = `<office:font-face-decls><style:font-face style:name="Liberation Serif" svg:font-family="'Liberation Serif'"/></office:font-face-decls>`;

function open() {
  const bytes = odsWith(SHEET, { automaticStyles: RICH_STYLE });
  // Put the font faces content.xml declares in front of its automatic styles.
  const { package: pkg } = readOdsWithPackage(bytes);
  const content = pkg.getText(OdsXmlNames.contentPart)!;
  pkg.setText(
    OdsXmlNames.contentPart,
    content.replace('<office:automatic-styles>', FONT_FACES + '<office:automatic-styles>'),
  );
  return readOdsWithPackage(pkg.write());
}

/** The automatic cell style the saved cell in `row` names. */
function savedStyle(bytes: Uint8Array, row: number): XmlElement {
  const root = parseXml(readEntry(bytes, OdsXmlNames.contentPart)).root;
  const rows = root.descendants(Table('table-row'));
  const name = rows[row]!.elements(Table('table-cell'))[0]!.attr(Table('style-name'))!;
  return root.descendants(Style('style')).find((s) => s.attr(Style('name')) === name)!;
}

const props = (style: XmlElement, local: string): XmlElement => style.element(Style(local))!;

describe('ODS cell styles keep what the model does not hold', () => {
  it('an open and save writes the source style back', () => {
    const { workbook, package: pkg } = open();
    const saved = writeOds(workbook, pkg);
    const style = savedStyle(saved, 0);
    expect(style.attr(Style('parent-style-name'))).toBe('Heading');
    const cell = props(style, 'table-cell-properties');
    expect(cell.attr(Style('diagonal-bl-tr'))).toBe('0.74pt solid #000000');
    expect(cell.attr(Style('shrink-to-fit'))).toBe('true');
    expect(cell.attr('urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', 'padding')).toBe('0.2cm');
    expect(cell.attr('urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', 'border-left')).toBe(
      '2.5pt solid #ff0000',
    );
    expect(cell.attr(Style('vertical-align'))).toBeNull();
    const paragraph = props(style, 'paragraph-properties');
    expect(paragraph.attr('urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', 'margin-left')).toBe(
      '0.5cm',
    );
    const text = props(style, 'text-properties');
    expect(text.attr(Style('text-underline-type'))).toBe('double');
    expect(text.attr('urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', 'font-variant')).toBe(
      'small-caps',
    );
    // The font it names is declared in the saved content.xml.
    expect(readEntry(saved, OdsXmlNames.contentPart)).toContain(
      'style:font-face style:name="Liberation Serif"',
    );
  });

  it('two styles the model reads alike stay apart', () => {
    const { workbook, package: pkg } = open();
    const saved = writeOds(workbook, pkg);
    expect(props(savedStyle(saved, 1), 'table-cell-properties').attr(Style('diagonal-tl-br'))).not.toBeNull();
    expect(props(savedStyle(saved, 2), 'table-cell-properties').attr(Style('diagonal-bl-tr'))).not.toBeNull();
  });

  it('making the cell bold keeps everything else', () => {
    const { workbook, package: pkg } = open();
    const sheet = workbook.activeSheet;
    new EditHistory(workbook).execute(
      EditCommands.applyStyle(workbook, sheet.name, CellRange.parse('A1'), (s) => ({
        ...s,
        font: { ...s.font, bold: true },
      })),
    );
    const style = savedStyle(writeOds(workbook, pkg), 0);
    expect(style.attr(Style('parent-style-name'))).toBe('Heading');
    const text = props(style, 'text-properties');
    expect(text.attr('urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', 'font-weight')).toBe(
      'bold',
    );
    expect(text.attr(Style('text-underline-type'))).toBe('double');
    expect(text.attr('urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', 'font-variant')).toBe(
      'small-caps',
    );
    expect(text.attr('urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', 'font-style')).toBe(
      'italic',
    );
    const cell = props(style, 'table-cell-properties');
    expect(cell.attr(Style('diagonal-bl-tr'))).toBe('0.74pt solid #000000');
    expect(cell.attr('urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', 'border-left')).toBe(
      '2.5pt solid #ff0000',
    );
    expect(cell.attr(Style('vertical-align'))).toBeNull();
    expect(
      props(style, 'paragraph-properties').attr(
        'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
        'margin-left',
      ),
    ).toBe('0.5cm');
  });

  it('removing italic a parent might supply writes it off explicitly', () => {
    const { workbook, package: pkg } = open();
    const sheet = workbook.activeSheet;
    new EditHistory(workbook).execute(
      EditCommands.applyStyle(workbook, sheet.name, CellRange.parse('A1'), (s) => ({
        ...s,
        font: { ...s.font, italic: false },
      })),
    );
    const text = props(savedStyle(writeOds(workbook, pkg), 0), 'text-properties');
    expect(text.attr('urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', 'font-style')).toBe(
      'normal',
    );
    expect(text.attr(Style('text-underline-type'))).toBe('double');
  });

  it('changing one border edge keeps the exact width of the others', () => {
    const { workbook, package: pkg } = readOdsWithPackage(
      odsWith(
        `<table:table table:name="Sheet1"><table:table-row><table:table-cell table:style-name="ce1"/></table:table-row></table:table>`,
        {
          automaticStyles: `<style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties fo:border="2.5pt solid #ff0000"/></style:style>`,
        },
      ),
    );
    const sheet = workbook.activeSheet;
    new EditHistory(workbook).execute(
      EditCommands.applyStyle(workbook, sheet.name, CellRange.parse('A1'), (s) => ({
        ...s,
        border: { ...s.border, top: BorderEdges.none },
      })),
    );
    const cell = props(savedStyle(writeOds(workbook, pkg), 0), 'table-cell-properties');
    const fo = 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0';
    expect(cell.attr(fo, 'border')).toBeNull();
    expect(cell.attr(fo, 'border-top')).toBe('none');
    for (const side of ['left', 'right', 'bottom'])
      expect(cell.attr(fo, `border-${side}`)).toBe('2.5pt solid #ff0000');
  });

  it('clear formatting drops the source style', () => {
    const { workbook, package: pkg } = open();
    const sheet = workbook.activeSheet;
    new EditHistory(workbook).execute(
      EditCommands.clearFormatting(workbook, sheet.name, CellRange.parse('A1')),
    );
    const root = parseXml(readEntry(writeOds(workbook, pkg), OdsXmlNames.contentPart)).root;
    const first = root.descendants(Table('table-row'))[0]!.elements(Table('table-cell'))[0]!;
    expect(first.attr(Table('style-name'))).toBeNull();
  });
});
