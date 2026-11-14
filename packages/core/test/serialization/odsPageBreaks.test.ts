import { describe, expect, it } from 'vitest';
import { readOds, readOdsWithPackage, writeOds } from '../../src/serialization/index.ts';
import { Defaults } from '../../src/defaults.ts';
import { odsWith, readEntry, roundTrip } from './helpers.ts';

/** Row 3 and column C start a new printed page, set by hand (Insert Page Break in LibreOffice). */
const STYLES = `
  <style:style style:name="co1" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="2.258cm"/></style:style>
  <style:style style:name="co2" style:family="table-column"><style:table-column-properties fo:break-before="page" style:column-width="2.258cm"/></style:style>
  <style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.452cm" fo:break-before="auto" style:use-optimal-row-height="true"/></style:style>
  <style:style style:name="ro2" style:family="table-row"><style:table-row-properties style:row-height="0.452cm" fo:break-before="page" style:use-optimal-row-height="true"/></style:style>`;

const TABLE = `<table:table table:name="S">
  <table:table-column table:style-name="co1" table:number-columns-repeated="2"/>
  <table:table-column table:style-name="co2"/>
  <table:table-column table:style-name="co1"/>
  <table:table-row table:style-name="ro1" table:number-rows-repeated="2"><table:table-cell/></table:table-row>
  <table:table-row table:style-name="ro2"><table:table-cell office:value-type="float" office:value="1"/></table:table-row>
  <table:table-row table:style-name="ro1"><table:table-cell/></table:table-row>
</table:table>`;

describe('Manual page breaks', () => {
  it('are read from row and column styles', () => {
    const sheet = readOds(odsWith(TABLE, { automaticStyles: STYLES })).findSheet('S')!;
    expect(sheet.findRow(2)?.pageBreak).toBe(true);
    expect(sheet.findRow(1)?.pageBreak ?? false).toBe(false);
    expect(sheet.findColumn(2)?.pageBreak).toBe(true);
    expect(sheet.findColumn(1)?.pageBreak ?? false).toBe(false);
    // An optimal-height row keeps the default height.
    expect(sheet.findRow(2)?.height).toBe(Defaults.rowHeight);
  });

  it('are written back as fo:break-before="page"', () => {
    const { workbook, package: pkg } = readOdsWithPackage(odsWith(TABLE, { automaticStyles: STYLES }));
    const content = readEntry(writeOds(workbook, pkg), 'content.xml');
    expect(content.match(/fo:break-before="page"/g)).toHaveLength(2);
    const sheet = roundTrip(workbook).findSheet('S')!;
    expect(sheet.findRow(2)?.pageBreak).toBe(true);
    expect(sheet.findColumn(2)?.pageBreak).toBe(true);
    expect(sheet.findRow(3)?.pageBreak ?? false).toBe(false);
    expect(sheet.findColumn(3)?.pageBreak ?? false).toBe(false);
  });

  it('keep a row height set alongside the break', () => {
    const { workbook } = readOdsWithPackage(odsWith(TABLE, { automaticStyles: STYLES }));
    workbook.findSheet('S')!.getOrCreateRow(2).height = 40;
    const row = roundTrip(workbook).findSheet('S')!.findRow(2)!;
    expect(row.pageBreak).toBe(true);
    expect(row.height).toBeCloseTo(40, 0);
  });
});
