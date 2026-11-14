import { describe, expect, it } from 'vitest';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { StructuralEditCommand } from '../../src/editing/structuralEditCommand.ts';
import { readOdsWithPackage, writeOds } from '../../src/serialization/index.ts';
import { tryParseXml, type XmlElement } from '../../src/xml/index.ts';
import { odsWith, readEntry } from './helpers.ts';

const cell = (v: number): string => `<table:table-cell office:value-type="float" office:value="${v}"/>`;
const row = (v: number, extra = ''): string => `<table:table-row${extra}>${cell(v)}</table:table-row>`;

/**
 * Column A is the print title column, B:C a column group. Row 1 is the print title row; rows 2–4
 * are a group holding a collapsed group of rows 3–4; rows 5–6 are outside any group.
 */
const TABLE = `<table:table table:name="S">
  <table:table-header-columns><table:table-column/></table:table-header-columns>
  <table:table-column-group><table:table-column table:number-columns-repeated="2"/></table:table-column-group>
  <table:table-column/>
  <table:table-header-rows>${row(1)}</table:table-header-rows>
  <table:table-row-group>
    ${row(2)}
    <table:table-row-group table:display="false">
      <table:table-row table:visibility="collapse" table:number-rows-repeated="2">${cell(3)}</table:table-row>
    </table:table-row-group>
  </table:table-row-group>
  ${row(5)}${row(6)}
</table:table>`;

function open() {
  const { workbook, package: pkg } = readOdsWithPackage(odsWith(TABLE));
  return { workbook, pkg, sheet: workbook.findSheet('S')! };
}

/** The saved table's grid as nested tag names, so the container structure can be compared. */
function grid(table: XmlElement): string[] {
  const out: string[] = [];
  const walk = (e: XmlElement, depth: number): void => {
    for (const child of e.elements()) {
      if (!child.local.startsWith('table-') || child.local === 'table-cell') continue;
      const repeat =
        child.attr({ uri: child.uri, local: 'number-rows-repeated', prefix: 'table' } as never) ?? '';
      const display =
        child.local.endsWith('group') && child.attributes.some((a) => a.name.local === 'display');
      out.push(
        `${'  '.repeat(depth)}${child.local}${repeat ? `×${repeat}` : ''}${display ? ' collapsed' : ''}`,
      );
      walk(child, depth + 1);
    }
  };
  walk(table, 0);
  return out;
}

function savedTable(bytes: Uint8Array): XmlElement {
  const doc = tryParseXml(readEntry(bytes, 'content.xml'))!;
  const body = doc.root.elements().find((e) => e.local === 'body')!;
  return body
    .elements()[0]!
    .elements()
    .find((e) => e.local === 'table')!;
}

describe('Row and column groups and print titles', () => {
  it('are read from their containers', () => {
    const { sheet } = open();
    expect(sheet.outline).toEqual({
      rowGroups: [
        { first: 1, last: 3, collapsed: false },
        { first: 2, last: 3, collapsed: true },
      ],
      columnGroups: [{ first: 1, last: 2, collapsed: false }],
      headerRows: { first: 0, last: 0 },
      headerColumns: { first: 0, last: 0 },
    });
    expect(sheet.getValue(4, 0)).toEqual({ kind: 'number', value: 5 });
  });

  it('are written back nested as they were', () => {
    const { workbook, pkg } = open();
    expect(grid(savedTable(writeOds(workbook, pkg)))).toEqual([
      'table-header-columns',
      '  table-column',
      'table-column-group',
      '  table-column',
      'table-header-rows',
      '  table-row',
      'table-row-group',
      '  table-row',
      '  table-row-group collapsed',
      '    table-row',
      '    table-row',
      'table-row',
      'table-row',
    ]);
    const reopened = readOdsWithPackage(writeOds(workbook, pkg)).workbook.findSheet('S')!;
    expect(reopened.outline).toEqual(open().sheet.outline);
  });

  it('do not fold repeated blank rows across a group edge', () => {
    const blank = '<table:table-row><table:table-cell/></table:table-row>';
    const { workbook, package: pkg } = readOdsWithPackage(
      odsWith(
        `<table:table table:name="S"><table:table-column/>${blank}<table:table-row-group>${blank}${blank}</table:table-row-group>${blank}${row(9)}</table:table>`,
      ),
    );
    const reopened = readOdsWithPackage(writeOds(workbook, pkg)).workbook.findSheet('S')!;
    expect(reopened.outline.rowGroups).toEqual([{ first: 1, last: 2, collapsed: false }]);
  });

  it('move with inserted and deleted rows and columns', () => {
    const { workbook, sheet } = open();
    const history = new EditHistory(workbook);
    history.execute(new StructuralEditCommand('S', 'row', 'insert', 0, 2));
    expect(sheet.outline.headerRows).toEqual({ first: 2, last: 2 });
    expect(sheet.outline.rowGroups[0]).toEqual({ first: 3, last: 5, collapsed: false });
    // Inserting inside a group grows it.
    history.execute(new StructuralEditCommand('S', 'row', 'insert', 4, 1));
    expect(sheet.outline.rowGroups[0]).toEqual({ first: 3, last: 6, collapsed: false });
    // Deleting a whole group removes it; the header row is untouched.
    history.execute(new StructuralEditCommand('S', 'row', 'delete', 5, 2));
    expect(sheet.outline.rowGroups).toEqual([{ first: 3, last: 4, collapsed: false }]);
    history.execute(new StructuralEditCommand('S', 'column', 'delete', 0, 1));
    expect(sheet.outline.headerColumns).toBeNull();
    expect(sheet.outline.columnGroups).toEqual([{ first: 0, last: 1, collapsed: false }]);
    // Undo puts every span back.
    for (let i = 0; i < 4; i++) history.undo();
    expect(sheet.outline).toEqual(open().sheet.outline);
  });
});
