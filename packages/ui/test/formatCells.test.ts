import { Argb, Borders, CellStyles, Fills, NumberFormats, type CellStyle } from '@cellestra/core';
import { describe, expect, it } from 'vitest';
import {
  buildFormatTransform,
  editFormatCells,
  firstVisibleEdge,
  formatSample,
  nextTriState,
  parseFontSize,
  presetIndex,
  seedFormatCells,
} from '../src/components/formatCells.ts';

const base = CellStyles.default;
const withFont = (patch: Partial<CellStyle['font']>): CellStyle => ({
  ...base,
  font: { ...base.font, ...patch },
});
const withCode = (code: string): CellStyle => ({ ...base, numberFormat: NumberFormats.of(code) });

describe('seedFormatCells', () => {
  it('marks fields the selection disagrees on as varies (null)', () => {
    const bold = withFont({ bold: true, sizePoints: 14 });
    const plain = base;
    const { values, varies } = seedFormatCells(bold, [bold, plain]);
    expect(values.bold).toBeNull();
    expect(values.size).toBeNull();
    expect(values.italic).toBe(false);
    expect(varies.bold).toBe(true);
    expect(varies.italic).toBe(false);
  });

  it('falls back to the active style when no selection styles are given', () => {
    const s = seedFormatCells(withFont({ italic: true }), []);
    expect(s.values.italic).toBe(true);
    expect(Object.values(s.varies).every((v) => !v)).toBe(true);
  });

  it('seeds a transparent text colour as black and the border picker from a visible edge', () => {
    const style: CellStyle = { ...withFont({ color: Argb.transparent }), border: Borders.all('medium') };
    const { values } = seedFormatCells(style, [style]);
    expect(values.textColor).toBe(Argb.black);
    expect(values.borderStyle).toBe('medium');
    expect(values.borderEdges).toEqual({ top: true, bottom: true, left: true, right: true });
    expect(firstVisibleEdge(Borders.none).style).toBe('thin');
  });
});

describe('buildFormatTransform', () => {
  it('leaves every cell alone when nothing is touched', () => {
    const bold = withFont({ bold: true, sizePoints: 14 });
    const t = buildFormatTransform(seedFormatCells(bold, [bold, base]));
    expect(t(bold)).toBe(bold);
    expect(t(base)).toBe(base);
  });

  it('keeps mixed number formats left on varies, and applies a chosen one to all', () => {
    const currency = withCode('$#,##0.00');
    const percent = withCode('0%');
    const seed = seedFormatCells(currency, [currency, percent]);
    expect(seed.values.numberFormat).toBeNull();
    expect(buildFormatTransform(seed)(percent).numberFormat.code).toBe('0%');
    const t = buildFormatTransform(editFormatCells(seed, { numberFormat: '0.00' }));
    expect(t(currency).numberFormat.code).toBe('0.00');
    expect(t(percent).numberFormat.code).toBe('0.00');
    expect(
      buildFormatTransform(editFormatCells(seed, { numberFormat: '  ' }))(percent).numberFormat.code,
    ).toBe('General');
  });

  it('rewrites only touched font fields', () => {
    const seed = seedFormatCells(base, [withFont({ italic: true }), base]);
    const t = buildFormatTransform(editFormatCells(seed, { bold: true, size: 18 }));
    const out = t(withFont({ italic: true }));
    expect(out.font).toMatchObject({ bold: true, italic: true, sizePoints: 18 });
    expect(t(base).font.italic).toBe(false);
  });

  it('applies a text colour once touched, even when the colours varied', () => {
    const red = withFont({ color: Argb.fromRgb(255, 0, 0) });
    const seed = seedFormatCells(red, [red, base]);
    expect(buildFormatTransform(seed)(red).font.color).toBe(red.font.color);
    const t = buildFormatTransform(editFormatCells(seed, { textColor: Argb.black }));
    expect(t(red).font.color).toBe(Argb.black);
  });

  it('turns fill on keeping each cell colour, off, or to a picked colour', () => {
    const yellow: CellStyle = { ...base, fill: Fills.solid(Argb.fromRgb(255, 255, 0)) };
    const seed = seedFormatCells(yellow, [yellow, base]);
    expect(seed.values.fill).toBeNull();
    const on = buildFormatTransform(editFormatCells(seed, { fill: true }));
    expect(on(yellow).fill.background).toBe(yellow.fill.background);
    expect(on(base).fill.background).toBe(yellow.fill.background); // fallback: the seeded colour
    const off = buildFormatTransform(editFormatCells(seed, { fill: false }));
    expect(off(yellow).fill.background).toBeNull();
    const blue = Argb.fromRgb(0, 0, 255);
    const picked = buildFormatTransform(editFormatCells(seed, { fill: true, fillColor: blue }));
    expect(picked(yellow).fill.background).toBe(blue);
  });

  it('sets alignment and wrap only when touched', () => {
    const centered: CellStyle = { ...base, alignment: { ...base.alignment, horizontal: 'center' } };
    const seed = seedFormatCells(centered, [centered, base]);
    expect(seed.values.horizontal).toBeNull();
    const t = buildFormatTransform(editFormatCells(seed, { vertical: 'top', wrapText: true }));
    expect(t(centered).alignment).toEqual({ horizontal: 'center', vertical: 'top', wrapText: true });
    expect(t(base).alignment.horizontal).toBe('general');
  });

  it('keeps a mixed border untouched, and writes the chosen edge to checked sides once touched', () => {
    const boxed: CellStyle = { ...base, border: Borders.all('thin') };
    const seed = seedFormatCells(boxed, [boxed, base]);
    const untouched = buildFormatTransform(seed);
    expect(Borders.hasAnyEdge(untouched(boxed).border)).toBe(true);
    expect(Borders.hasAnyEdge(untouched(base).border)).toBe(false);

    const red = Argb.fromRgb(255, 0, 0);
    const t = buildFormatTransform(
      editFormatCells(seed, {
        borderEdges: { top: false, bottom: true, left: false, right: false },
        borderStyle: 'double',
        borderColor: red,
      }),
    );
    const out = t(base).border;
    expect(out.bottom).toEqual({ style: 'double', color: red });
    expect(out.top.style).toBe('none');
    expect(t(boxed).border.left.style).toBe('none');
  });
});

describe('dialog helpers', () => {
  it('cycles tri-state checkboxes, returning to varies only for varying fields', () => {
    expect(nextTriState(null, true)).toBe(true);
    expect(nextTriState(true, true)).toBe(false);
    expect(nextTriState(false, true)).toBeNull();
    expect(nextTriState(false, false)).toBe(true);
  });

  it('un-touches a field set back to varies', () => {
    const bold = withFont({ bold: true });
    const seed = seedFormatCells(bold, [bold, base]);
    const back = editFormatCells(editFormatCells(seed, { bold: true }), { bold: null });
    expect(back.touched.has('bold')).toBe(false);
    expect(buildFormatTransform(back)(base)).toBe(base);
  });

  it('clamps font sizes and finds presets', () => {
    expect(parseFontSize('200')).toBe(96);
    expect(parseFontSize('2')).toBe(6);
    expect(parseFontSize('abc')).toBeNull();
    expect(presetIndex('general')).toBe(0);
    expect(presetIndex('0.00%')).toBeGreaterThan(0);
    expect(presetIndex('0.000')).toBe(-1);
  });

  it('formats a live sample for the code', () => {
    expect(formatSample('#,##0.00', 'en-US')).toEqual({ text: '1,234.57', negative: '-1,234.57' });
    expect(formatSample('0%', 'en-US').text).toBe('12%');
    expect(formatSample('yyyy-mm-dd', 'en-US').text).toBe('2025-03-14');
    expect(formatSample('#,##0.00', 'de-DE').text).toBe('1.234,57');
  });
});
