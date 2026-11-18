import { describe, expect, it } from 'vitest';
import { CellRange, Workbook, table } from '@cellestra/core';
import { Axis } from '../src/grid/axis.ts';
import { chromeButtons } from '../src/grid/chrome.ts';
import { HEADER_HEIGHT, HEADER_WIDTH, Viewport } from '../src/grid/viewport.ts';

const viewport = (scrollX: number) =>
  new Viewport({
    columns: new Axis(16384, 64, []),
    rows: new Axis(1048576, 20, []),
    frozenColumns: 1,
    frozenRows: 0,
    scrollX,
    scrollY: 0,
    zoom: 1,
    width: HEADER_WIDTH + 64 * 6,
    height: HEADER_HEIGHT + 20 * 10,
  });

describe('chromeButtons with frozen columns', () => {
  const sheet = Workbook.createNew().activeSheet;
  sheet.addTable(table('T', CellRange.parse('A1:E10')));
  const columns = (scrollX: number) =>
    chromeButtons(sheet, viewport(scrollX)).map((b) => (b.kind === 'filter' ? b.column : -1));

  it('shows a chevron for every header cell when nothing is scrolled', () => {
    expect(columns(0)).toEqual([0, 1, 2, 3, 4]);
  });

  it('drops the chevron of a column scrolled under the frozen pane', () => {
    // Column B starts 55px under column A: its chevron would sit over A's header cell.
    const v = viewport(55);
    const shown = chromeButtons(sheet, v);
    expect(shown.map((b) => (b.kind === 'filter' ? b.column : -1))).toEqual([0, 2, 3, 4]);
    for (const b of shown)
      if (b.kind === 'filter' && b.column > 0) expect(b.rect.x).toBeGreaterThanOrEqual(v.paneX);
  });
});
