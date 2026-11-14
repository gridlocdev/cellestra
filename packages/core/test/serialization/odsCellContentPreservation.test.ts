import { describe, expect, it } from 'vitest';
import { DataEditCommands } from '../../src/editing/dataEditCommands.ts';
import { EditCommands } from '../../src/editing/editCommands.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { StructuralEditCommand } from '../../src/editing/structuralEditCommand.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { OdsXmlNames, readOdsWithPackage, writeOds } from '../../src/serialization/index.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { odsWith, readEntry } from './helpers.ts';

const NOTE = '<office:annotation><dc:creator>Ann</dc:creator><text:p>Check this</text:p></office:annotation>';
const FRAME =
  '<draw:frame draw:name="Logo" svg:width="2cm" svg:height="1cm" svg:x="0cm" svg:y="0cm"><draw:image xlink:href="Pictures/logo.png"/></draw:frame>';

/** A1 holds 3 with a comment, B1 holds 1 with an anchored image, A2 is a comment and nothing else. */
const SHEET = `<table:table table:name="Sheet1"><table:table-column table:number-columns-repeated="2"/>
  <table:table-row><table:table-cell office:value-type="float" office:value="3">${NOTE}<text:p>3</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="1">${FRAME}<text:p>1</text:p></table:table-cell></table:table-row>
  <table:table-row><table:table-cell>${NOTE.replace('Check this', 'Only a note')}</table:table-cell></table:table-row>
</table:table>`;

function open() {
  const { workbook, package: pkg } = readOdsWithPackage(odsWith(SHEET));
  return { workbook, pkg, sheet: workbook.activeSheet, history: new EditHistory(workbook) };
}

/** The saved content.xml, reopened so the cell children can be checked by position. */
function cellXml(workbook: ReturnType<typeof open>['workbook'], pkg: ReturnType<typeof open>['pkg']): string {
  return readEntry(writeOds(workbook, pkg), OdsXmlNames.contentPart);
}

describe('ODS cell comments and anchored shapes', () => {
  it('survive inserting a row above them', () => {
    const { workbook, pkg, sheet, history } = open();
    history.execute(new StructuralEditCommand(sheet.name, 'row', 'insert', 0, 1));
    const content = cellXml(workbook, pkg);
    expect(content).toContain('Check this');
    expect(content).toContain('Only a note');
    expect(content).toContain('draw:name="Logo"');
    const reopened = readOdsWithPackage(writeOds(workbook, pkg)).workbook.activeSheet;
    expect(reopened.getValue(1, 0)).toEqual(CellValue.number(3));
  });

  it('survive deleting a column before them and undoing it', () => {
    const { workbook, pkg, sheet, history } = open();
    history.execute(new StructuralEditCommand(sheet.name, 'column', 'insert', 0, 2));
    history.undo();
    const content = cellXml(workbook, pkg);
    expect(content).toContain('Check this');
    expect(content).toContain('draw:name="Logo"');
    expect(content).toContain('Only a note');
  });

  it('travel with their cell when it is sorted', () => {
    const { workbook, pkg, sheet, history } = open();
    sheet.setValue(1, 0, CellValue.number(2));
    history.execute(DataEditCommands.sortByColumn(sheet.name, CellRange.parse('A1:A2'), 0, true, false));
    expect(sheet.getValue(1, 0)).toEqual(CellValue.number(3));
    const content = cellXml(workbook, pkg);
    expect(content).toContain('Check this');
    // The note follows the 3 down to the second row.
    expect(content.indexOf('Check this')).toBeGreaterThan(content.indexOf('office:value="2"'));
  });

  it('a comment stays when the cell value is deleted, and undo brings back a comment-only cell', () => {
    const { workbook, pkg, sheet, history } = open();
    history.execute(EditCommands.clearContents(sheet.name, CellRange.parse('A1:A2')));
    expect(sheet.getValue(0, 0)).toEqual(CellValue.empty);
    expect(cellXml(workbook, pkg)).toContain('Check this');

    history.execute(EditCommands.clearAll(sheet.name, CellRange.parse('A2')));
    expect(cellXml(workbook, pkg)).not.toContain('Only a note');
    history.undo();
    expect(cellXml(workbook, pkg)).toContain('Only a note');
  });
});
