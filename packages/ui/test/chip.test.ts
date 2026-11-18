import { describe, expect, it } from 'vitest';
import {
  Argb,
  CellRange,
  CellValue,
  CHIP_COLORS,
  chipColorsTransform,
  ConditionalStyles,
  dataValidation,
  Workbook,
} from '@cellestra/core';
import { Axis } from '../src/grid/axis.ts';
import { chipColors, chipRect, chevronRect } from '../src/grid/chip.ts';
import { chromeButtons } from '../src/grid/chrome.ts';
import { darkPalette, lightPalette } from '../src/grid/palette.ts';
import { renderGrid } from '../src/grid/render.ts';
import { HEADER_HEIGHT, HEADER_WIDTH, Viewport } from '../src/grid/viewport.ts';

const red = CHIP_COLORS[0]!;
const viewport = () =>
  new Viewport({
    columns: new Axis(16384, 64, []),
    rows: new Axis(1048576, 20, []),
    frozenColumns: 0,
    frozenRows: 0,
    scrollX: 0,
    scrollY: 0,
    zoom: 1,
    width: HEADER_WIDTH + 64 * 6,
    height: HEADER_HEIGHT + 20 * 10,
  });

/** A workbook whose B2:B4 is a Yes/No list with Yes in red; B2 holds Yes, B3 No, B4 is empty. */
function listSheet() {
  const wb = Workbook.createNew();
  const sheet = wb.activeSheet;
  const target = CellRange.parse('B2:B4');
  sheet.addValidation(dataValidation('list', target, { listValues: ['Yes', 'No'] }));
  const style = chipColorsTransform([{ value: 'Yes', color: red }], [], {
    sheet: sheet.name,
    row: 1,
    column: 1,
  })(wb.styles.get(0));
  const id = wb.styles.intern(style);
  for (let row = 1; row <= 3; row++) sheet.getOrCreate(row, 1).styleId = id;
  sheet.setValue(1, 1, CellValue.text('Yes'));
  sheet.setValue(2, 1, CellValue.text('No'));
  return wb;
}

describe('list chips', () => {
  it('puts a chevron button on every visible list cell, inside its chip', () => {
    const wb = listSheet();
    const buttons = chromeButtons(wb.activeSheet, viewport()).filter((b) => b.kind === 'validation');
    expect(buttons.map((b) => (b.kind === 'validation' ? [b.row, b.column] : []))).toEqual([
      [1, 1],
      [2, 1],
      [3, 1],
    ]);
    const cell = viewport().rangeRect(1, 1, 1, 1);
    const chip = chipRect(cell, 1);
    const b = buttons[0]!.rect;
    expect(b.x + b.w).toBe(chip.x + chip.w);
    expect(b.x).toBeGreaterThan(cell.x + cell.w / 2);
  });

  it('draws an item in its colour on a light theme and swaps the colours on a dark one', () => {
    expect(chipColors(red, lightPalette)).toEqual({ fill: '#FFC7CE', text: '#9C0006' });
    const dark = chipColors(red, darkPalette);
    expect(dark.text).toBe('#FFC7CE');
    expect(dark.fill).toMatch(/^rgb\(1\d\d,0,4\)$/);
    expect(chipColors(null, darkPalette)).toEqual({ fill: darkPalette.chipFill, text: darkPalette.chipText });
  });

  it('colours the chip, not the cell, and gives an empty cell only a chevron', () => {
    const wb = listSheet();
    const fills: { kind: 'rect' | 'chip'; color: unknown }[] = [];
    const state: Record<string | symbol, unknown> = {};
    const ctx = new Proxy(state, {
      get(target, key) {
        if (key in target) return target[key];
        if (key === 'measureText') return (text: string) => ({ width: text.length * 6 });
        if (key === 'fillRect') return () => fills.push({ kind: 'rect', color: target.fillStyle });
        if (key === 'fill') return () => fills.push({ kind: 'chip', color: target.fillStyle });
        return () => {};
      },
      set(target, key, value) {
        target[key] = value;
        return true;
      },
    }) as unknown as CanvasRenderingContext2D;
    const conditional = new ConditionalStyles(wb);
    renderGrid({
      ctx,
      viewport: viewport(),
      sheet: wb.activeSheet,
      styles: wb.styles,
      palette: lightPalette,
      selection: { row: 0, column: 0, ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }] },
      display: (v) => ({ text: v.kind === 'text' ? v.value : '' }),
      conditional: (style, row, column) => conditional.resolve(style, wb.activeSheet.name, row, column),
    });
    const redCss = Argb.toCss(red.fill);
    expect(fills.filter((f) => f.kind === 'rect' && f.color === redCss)).toEqual([]);
    expect(fills.filter((f) => f.kind === 'chip').map((f) => f.color)).toEqual([
      redCss,
      lightPalette.chipFill,
    ]);
  });

  it('keeps the chevron clear of a narrow cell', () => {
    const r = chevronRect({ x: 0, y: 0, w: 14, h: 20 }, 1);
    expect(r.w).toBeLessThanOrEqual(6);
  });
});
