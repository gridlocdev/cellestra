import { describe, expect, it } from 'vitest';
import { CellRange, CellStyles, DataEditCommands, Workbook, table } from '@cellestra/core';
import { Axis } from '../src/grid/axis.ts';
import { lightPalette } from '../src/grid/palette.ts';
import { checkboxRect, fillHandleRect, renderGrid } from '../src/grid/render.ts';
import { HEADER_HEIGHT, HEADER_WIDTH, Viewport } from '../src/grid/viewport.ts';

/** A canvas context that records strokeRect calls (with the stroke colour and width at the time). */
function recordingContext() {
  const strokes: { rect: number[]; color: unknown; width: unknown }[] = [];
  const state: Record<string | symbol, unknown> = {};
  const ctx = new Proxy(state, {
    get(target, key) {
      if (key in target) return target[key];
      if (key === 'measureText') return (text: string) => ({ width: text.length * 6 });
      if (key === 'strokeRect')
        return (...rect: number[]) =>
          strokes.push({ rect, color: target.strokeStyle, width: target.lineWidth });
      return () => {};
    },
    set(target, key, value) {
      target[key] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, strokes };
}

function draw(workbook: Workbook) {
  const sheet = workbook.worksheets[0]!;
  const { ctx, strokes } = recordingContext();
  renderGrid({
    ctx,
    viewport: new Viewport({
      columns: new Axis(16384, 64, []),
      rows: new Axis(1048576, 20, []),
      frozenColumns: 0,
      frozenRows: 0,
      scrollX: 0,
      scrollY: 0,
      zoom: 1,
      width: HEADER_WIDTH + 64 * 10,
      height: HEADER_HEIGHT + 20 * 20,
    }),
    sheet,
    styles: workbook.styles,
    palette: lightPalette,
    selection: { row: 10, column: 8, ranges: [{ top: 10, left: 8, bottom: 10, right: 8 }] },
    display: () => ({ text: '' }),
  });
  return strokes.filter((s) => s.color === lightPalette.tableOutline);
}

describe('table outline', () => {
  it('frames the table with a line heavier than the gridlines', () => {
    const wb = Workbook.createNew();
    wb.worksheets[0]!.addTable(table('Data', CellRange.parse('A1:B3')));
    expect(draw(wb)).toEqual([
      { rect: [HEADER_WIDTH, HEADER_HEIGHT, 128, 60], color: lightPalette.tableOutline, width: 2 },
    ]);
  });

  it('grows with the table when an entry extends it', () => {
    const wb = Workbook.createNew();
    const sheet = wb.worksheets[0]!;
    const t = table('Data', CellRange.parse('A1:B3'));
    sheet.addTable(t);
    // Typing just below the table (A4) grows it to A1:B4.
    DataEditCommands.expandTableForEdit(wb, sheet.name, t, 3, 0)!.apply(wb);
    expect(draw(wb).map((s) => s.rect)).toEqual([[HEADER_WIDTH, HEADER_HEIGHT, 128, 80]]);
    // And just right of it (C1) to A1:C4.
    DataEditCommands.expandTableForEdit(wb, sheet.name, sheet.tables[0]!, 0, 2)!.apply(wb);
    expect(draw(wb).map((s) => s.rect)).toEqual([[HEADER_WIDTH, HEADER_HEIGHT, 192, 80]]);
  });
});

describe('fillHandleRect', () => {
  const viewport = () =>
    new Viewport({
      columns: new Axis(16384, 64, []),
      rows: new Axis(1048576, 20, []),
      frozenColumns: 0,
      frozenRows: 0,
      scrollX: 0,
      scrollY: 0,
      zoom: 1,
      width: 800,
      height: 600,
    });

  it('sits on the bottom-right corner of the last range', () => {
    const v = viewport();
    const r = fillHandleRect(v, { row: 0, column: 0, ranges: [{ top: 0, left: 0, bottom: 1, right: 1 }] })!;
    expect(r.x + 4).toBe(HEADER_WIDTH + 128);
    expect(r.y + 4).toBe(HEADER_HEIGHT + 40);
  });

  it('is absent for whole rows and columns', () => {
    const v = viewport();
    expect(
      fillHandleRect(v, { row: 0, column: 0, ranges: [{ top: 0, left: 2, bottom: 1048575, right: 2 }] }),
    ).toBeNull();
    expect(
      fillHandleRect(v, { row: 0, column: 0, ranges: [{ top: 3, left: 0, bottom: 3, right: 16383 }] }),
    ).toBeNull();
  });
});

describe('checkboxRect', () => {
  const cell = { x: 100, y: 40, w: 64, h: 20 };
  const aligned = (horizontal: 'general' | 'left' | 'right', vertical: 'top' | 'middle' | 'bottom') => ({
    ...CellStyles.default,
    alignment: { ...CellStyles.default.alignment, horizontal, vertical },
  });

  it('centres a 13px box in a default cell', () => {
    expect(checkboxRect(cell, CellStyles.default, 1)).toEqual({ x: 126 + 0.5, y: 44 + 0.5, w: 13, h: 13 });
  });

  it('follows the alignment and shrinks to fit', () => {
    expect(checkboxRect(cell, aligned('left', 'top'), 1).x).toBeLessThan(110);
    expect(checkboxRect(cell, aligned('right', 'middle'), 1).x).toBeGreaterThan(140);
    expect(checkboxRect({ ...cell, h: 12 }, CellStyles.default, 1).h).toBe(8);
    expect(checkboxRect(cell, CellStyles.default, 2).w).toBe(16);
  });
});
