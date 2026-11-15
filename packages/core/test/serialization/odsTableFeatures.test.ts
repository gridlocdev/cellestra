import { describe, expect, it } from 'vitest';
import { CellRange } from '../../src/model/cellRange.ts';
import { MergedRegion } from '../../src/model/mergedRegion.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { readOds, writeOds } from '../../src/serialization/index.ts';
import { runFormat, type StyledRun } from '../../src/serialization/ods/odsCellStyleRuns.ts';
import { addTableForms } from '../../src/serialization/ods/odsCheckboxes.ts';
import { readRowVisibility } from '../../src/serialization/ods/odsHiddenLines.ts';
import { splitCellChildren } from '../../src/serialization/ods/odsPreserved.ts';
import { Office, Table, Text } from '../../src/serialization/ods/odsXmlNames.ts';
import { CellStyles, Fonts } from '../../src/styling/cellStyle.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { XmlElement } from '../../src/xml/index.ts';
import { odsWith, readEntry, roundTrip } from './helpers.ts';

function boldStyle(workbook: Workbook): number {
  return workbook.styles.intern({ ...CellStyles.default, font: { ...Fonts.default, bold: true } });
}

describe('rows and columns', () => {
  it('keep their width, height, page break and whole-line format through a save', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    const bold = boldStyle(workbook);
    sheet.getOrCreateColumn(3).width = 140;
    sheet.getOrCreateColumn(3).pageBreak = true;
    sheet.getOrCreateColumn(5).styleId = bold;
    sheet.getOrCreateRow(2).height = 44;
    sheet.getOrCreateRow(7).styleId = bold;

    const back = roundTrip(workbook).activeSheet;
    expect(Math.abs(back.getColumnWidth(3) - 140)).toBeLessThanOrEqual(0.5);
    expect(back.findColumn(3)?.pageBreak).toBe(true);
    expect(back.findColumn(5)?.styleId).not.toBe(back.findColumn(4)?.styleId ?? 0);
    expect(Math.abs(back.getRowHeight(2) - 44)).toBeLessThanOrEqual(0.5);
    expect(back.findRow(7)?.styleId).toBe(back.findColumn(5)?.styleId);
  });
});

describe('hidden and filtered lines', () => {
  it('keep a filter-hidden row apart from one hidden by hand', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.getOrCreateRow(1).filtered = true;
    sheet.getOrCreateRow(2).manuallyHidden = true;
    sheet.getOrCreateColumn(4).hidden = true;

    const back = roundTrip(workbook).activeSheet;
    expect(back.findRow(1)).toMatchObject({ filtered: true, manuallyHidden: false });
    expect(back.findRow(2)).toMatchObject({ filtered: false, manuallyHidden: true });
    expect(back.isColumnHidden(4)).toBe(true);
  });

  it('writes a row both filtered and hidden by hand as filtered', () => {
    const workbook = Workbook.createNew();
    const row = workbook.activeSheet.getOrCreateRow(0);
    row.filtered = true;
    row.manuallyHidden = true;
    const content = readEntry(writeOds(workbook), 'content.xml');
    expect(content).toContain('table:visibility="filter"');
    expect(content).not.toContain('table:visibility="collapse"');
  });

  it('reads `collapse` as hidden by hand and `filter` as filtered', () => {
    const row = (visibility: string) =>
      new XmlElement(Table('table-row')).setAttr(Table('visibility'), visibility);
    expect(readRowVisibility(row('collapse'))).toEqual({ filtered: false, manuallyHidden: true });
    expect(readRowVisibility(row('filter'))).toEqual({ filtered: true, manuallyHidden: false });
    expect(readRowVisibility(new XmlElement(Table('table-row')))).toEqual({
      filtered: false,
      manuallyHidden: false,
    });
  });
});

describe('merged cells', () => {
  it('survive a save, with the covered cell keeping its content', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.addMerge(new MergedRegion(CellRange.parse('B2:D3')));
    sheet.setValue(1, 1, CellValue.text('anchor'));
    sheet.setValue(2, 3, CellValue.text('under'));

    const back = roundTrip(workbook).activeSheet;
    expect(back.merges.map((m) => m.range.toString())).toEqual(['B2:D3']);
    expect(back.getValue(2, 3)).toEqual(CellValue.text('under'));
  });

  it('read a repeated spanned cell as side-by-side merges, and drop a later overlapping span', () => {
    const bytes = odsWith(
      `<table:table table:name="S">
        <table:table-row><table:table-cell table:number-columns-spanned="2" table:number-columns-repeated="4"/></table:table-row>
        <table:table-row><table:table-cell table:number-rows-spanned="2"/></table:table-row>
        <table:table-row><table:table-cell table:number-columns-spanned="2"/></table:table-row>
      </table:table>`,
    );
    const sheet = readOds(bytes).activeSheet;
    expect(sheet.merges.map((m) => m.range.toString())).toEqual(['A1:B1', 'C1:D1', 'A2:A3']);
  });
});

describe('per-cell styles', () => {
  const run = (over: Partial<StyledRun>): StyledRun => ({
    column: 0,
    repeat: 1,
    rowCount: 1,
    hasContent: false,
    styleId: 3,
    lastInRow: false,
    rowStyleBefore: undefined,
    trailingRows: false,
    ...over,
  });

  it('read a styled blank run to the end of the row as a row format', () => {
    expect(runFormat(run({ repeat: 16384, rowCount: 5, lastInRow: true }))).toBe('row');
  });

  it('read a big styled blank run down to the end of the sheet as a column format', () => {
    expect(runFormat(run({ repeat: 2, rowCount: 1 << 20, trailingRows: true }))).toBe('column');
  });

  it('keep a small run, or one with content, as cells', () => {
    expect(runFormat(run({ repeat: 3, rowCount: 2, lastInRow: true }))).toBeNull();
    expect(runFormat(run({ repeat: 16384, rowCount: 5, lastInRow: true, hasContent: true }))).toBeNull();
  });

  it('survive a save, and the default style is left unnamed', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.text('plain'));
    sheet.setValue(0, 1, CellValue.text('bold'));
    sheet.getOrCreate(0, 1).styleId = boldStyle(workbook);
    const bytes = writeOds(workbook);
    const content = readEntry(bytes, 'content.xml');
    expect(content).toMatch(/<table:table-cell office:value-type="string"/);
    const back = readOds(bytes);
    expect(back.styles.get(back.activeSheet.find(0, 1)!.styleId).font.bold).toBe(true);
  });
});

describe('preserved content', () => {
  it('keeps unmodeled table attributes and a cell comment on its side of the text', () => {
    const bytes = odsWith(
      `<table:table table:name="S" table:protected="true">
        <table:table-row><table:table-cell office:value-type="string"><office:annotation><text:p>note</text:p></office:annotation><text:p>x</text:p></table:table-cell></table:table-row>
      </table:table>`,
    );
    const content = readEntry(writeOds(readOds(bytes)), 'content.xml');
    expect(content).toContain('table:protected="true"');
    expect(content).toMatch(/<office:annotation>.*note.*<\/office:annotation><text:p>x<\/text:p>/s);
  });

  it('splits a cell into its paragraphs and kept children, dropping what the model takes', () => {
    const cell = new XmlElement(Table('table-cell')).add(
      new XmlElement(Office('annotation')),
      new XmlElement(Text('p')),
      new XmlElement(Office('forms')),
      new XmlElement(Table('detective')),
    );
    const { paragraphs, preserved } = splitCellChildren(cell, (c) => c.is(Office('forms')));
    expect(paragraphs).toHaveLength(1);
    expect(preserved?.before.map((e) => e.local)).toEqual(['annotation']);
    expect(preserved?.after.map((e) => e.local)).toEqual(['detective']);
  });
});

describe('checkbox forms in a table', () => {
  const form = () => new XmlElement(Office('forms'));
  const names = (table: XmlElement) => table.elements().map((e) => e.local);

  it('puts a new office:forms after the title and before the shapes', () => {
    const table = new XmlElement(Table('table'));
    const checkboxes = new XmlElement(Table('checkbox-form-stand-in'));
    addTableForms(table, [new XmlElement(Table('title')), new XmlElement(Table('shapes'))], checkboxes);
    expect(names(table)).toEqual(['title', 'forms', 'shapes']);
  });

  it("joins the file's own office:forms", () => {
    const table = new XmlElement(Table('table'));
    const checkboxes = new XmlElement(Table('checkbox-form-stand-in'));
    addTableForms(table, [form()], checkboxes);
    expect(names(table)).toEqual(['forms']);
    expect(
      table
        .elements()[0]!
        .elements()
        .map((e) => e.local),
    ).toEqual(['checkbox-form-stand-in']);
  });
});
