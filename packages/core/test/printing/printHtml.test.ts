import {
  CellRange,
  CellValue,
  MergedRegion,
  PageSetups,
  SheetDimensions,
  Workbook,
  type PageSetup,
} from '@cellestra/core';
import { describe, expect, it } from 'vitest';
import { printHtml, printableRange, usedRange } from '../../src/printing/printHtml.ts';

const display = (v: CellValue) => (v.kind === 'number' ? String(v.value) : v.kind === 'text' ? v.value : '');

function grid(rows: number, columns: number) {
  const wb = Workbook.createNew();
  const sheet = wb.worksheets[0]!;
  for (let r = 0; r < rows; r++)
    for (let c = 0; c < columns; c++) sheet.setValue(r, c, CellValue.number(r * 100 + c));
  return { wb, sheet };
}

const print = (rows: number, columns: number, setup: PageSetup | null) => {
  const { wb, sheet } = grid(rows, columns);
  return printHtml({ sheet, styles: wb.styles, setup, range: null, display, title: 'T' });
};

describe('printHtml', () => {
  it('prints a small sheet on one page with the paper size in @page', () => {
    const doc = print(3, 3, null);
    expect(doc.pages).toBe(1);
    expect(doc.html).toContain('size:210mm 297mm');
    expect(doc.html).toContain('>202<');
  });

  it('splits a wide sheet across pages at 100%', () => {
    expect(print(10, 30, null).pages).toBeGreaterThan(1);
  });

  it('fits a wide sheet to one page across', () => {
    const doc = print(10, 30, { ...PageSetups.default, scale: { kind: 'fit', width: 1, height: 0 } });
    expect(doc.pages).toBe(1);
  });

  it('fits on a page count', () => {
    const doc = print(200, 30, { ...PageSetups.default, scale: { kind: 'pages', pages: 2 } });
    expect(doc.pages).toBeLessThanOrEqual(2);
  });

  it('prints row and column headers when asked', () => {
    const doc = print(2, 2, { ...PageSetups.default, headers: true });
    expect(doc.html).toMatch(/>B</);
  });

  it('shows a merge once, with its span', () => {
    const { wb, sheet } = grid(0, 0);
    sheet.setValue(0, 0, CellValue.text('Merged'));
    sheet.addMerge(new MergedRegion(CellRange.parse('A1:C2')));
    const doc = printHtml({ sheet, styles: wb.styles, setup: null, range: null, display, title: 'T' });
    expect(doc.html.match(/Merged/g)).toHaveLength(1);
    expect(doc.html).toMatch(/colspan="3"/);
    expect(doc.html).toMatch(/rowspan="2"/);
  });

  it('escapes cell text', () => {
    const { wb, sheet } = grid(0, 0);
    sheet.setValue(0, 0, CellValue.text('<b>&'));
    const doc = printHtml({ sheet, styles: wb.styles, setup: null, range: null, display, title: 'T' });
    expect(doc.html).toContain('&lt;b&gt;&amp;');
  });

  it('finds the used range', () => {
    const { sheet } = grid(4, 6);
    expect(usedRange(sheet)).toEqual({ top: 0, left: 0, bottom: 3, right: 5 });
  });
});

describe('printHtml row heights', () => {
  it('paginates and sizes rows by the height their content needs', () => {
    const { wb, sheet } = grid(40, 2);
    const plain = printHtml({ sheet, styles: wb.styles, setup: null, range: null, display, title: 'T' });
    expect(plain.pages).toBe(1);
    // Every row wraps to three lines: 40 of them no longer fit on one page.
    const contentHeights = new Map(Array.from({ length: 40 }, (_, r) => [r, 60] as [number, number]));
    const tall = printHtml({
      sheet,
      styles: wb.styles,
      setup: null,
      range: null,
      display,
      title: 'T',
      contentHeights,
    });
    expect(tall.pages).toBeGreaterThan(1);
    expect(tall.html).toContain('<tr style="height:60px">');
  });
});

describe('printing a selection', () => {
  const lastRow = SheetDimensions.maxRows - 1;
  const lastColumn = SheetDimensions.maxColumns - 1;

  it('prints only the used part of whole columns, rows or the whole sheet', () => {
    const { wb, sheet } = grid(3, 3);
    const all = { top: 0, left: 0, bottom: lastRow, right: lastColumn };
    expect(printableRange(sheet, all)).toEqual({ top: 0, left: 0, bottom: 2, right: 2 });
    expect(printableRange(sheet, { top: 0, left: 1, bottom: lastRow, right: 1 })).toEqual({
      top: 0,
      left: 1,
      bottom: 2,
      right: 1,
    });
    const range = printableRange(sheet, all);
    const doc = printHtml({ sheet, styles: wb.styles, setup: null, range, display, title: 'T' });
    expect(doc.pages).toBe(1);
    expect(doc.html).toContain('>202<');
  });

  it('keeps a merge that reaches past the last value', () => {
    const { sheet } = grid(2, 2);
    sheet.addMerge(new MergedRegion(CellRange.parse('B2:D5')));
    expect(printableRange(sheet, { top: 0, left: 0, bottom: lastRow, right: lastColumn })).toEqual({
      top: 0,
      left: 0,
      bottom: 4,
      right: 3,
    });
  });

  it('has nothing to print for a selection wholly outside the used cells', () => {
    const { sheet } = grid(3, 3);
    expect(printableRange(sheet, { top: 10, left: 0, bottom: lastRow, right: 0 })).toBeNull();
    expect(printableRange(sheet, { top: 0, left: 5, bottom: 1, right: 6 })).toBeNull();
  });
});
