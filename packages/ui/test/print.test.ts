import { CellRange, CellValue, MergedRegion, PageSetups, Workbook, type PageSetup } from '@cellestra/core';
import { describe, expect, it } from 'vitest';
import { draftOf, setupOf } from '../src/components/pageSetupForm.ts';
import { printHtml, usedRange } from '../src/print/printHtml.ts';

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

describe('page setup form', () => {
  it('round-trips a setup through the draft', () => {
    const setup: PageSetup = {
      ...PageSetups.orient(PageSetups.default, 'landscape'),
      scale: { kind: 'fit', width: 1, height: 0 },
      gridlines: true,
    };
    const back = setupOf(draftOf(setup), setup);
    expect('setup' in back && PageSetups.equal(back.setup, setup)).toBe(true);
  });

  it('turns the paper with the orientation', () => {
    const r = setupOf({ ...draftOf(null), paper: 'letter', orientation: 'landscape' }, null);
    expect('setup' in r && [r.setup.paperWidth, r.setup.paperHeight]).toEqual([279.4, 215.9]);
  });

  it('rejects bad values', () => {
    const d = draftOf(null);
    expect(setupOf({ ...d, top: 'x' }, null)).toHaveProperty('error');
    expect(setupOf({ ...d, left: '100', right: '100' }, null)).toHaveProperty('error');
    expect(setupOf({ ...d, percent: '5' }, null)).toHaveProperty('error');
    expect(setupOf({ ...d, scaleKind: 'fit', fitWidth: '', fitHeight: '' }, null)).toHaveProperty('error');
    expect(setupOf({ ...d, scaleKind: 'pages', pages: '0' }, null)).toHaveProperty('error');
  });
});
