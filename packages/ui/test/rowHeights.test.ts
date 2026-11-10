import { describe, expect, it } from 'vitest';
import { CellRange, CellValue, Defaults, MergedRegion, Workbook, type CellStyle } from '@cellestra/core';
import { ContentHeights, contentRowHeights, textHeight, type RowMeasure } from '../src/grid/rowHeights.ts';

/** Every character is 6px wide; wrapping is by character count. */
const measure: RowMeasure = {
  display: (v) => (v.kind === 'text' ? v.value : String(v.kind === 'number' ? v.value : '')),
  wrappedLines: (text, _style, width) => Math.max(1, Math.ceil((text.length * 6) / width)),
  columnWidth: () => 64,
};

function styled(change: (s: CellStyle) => CellStyle) {
  const wb = Workbook.createNew();
  const sheet = wb.worksheets[0]!;
  const id = wb.styles.mutate(0, change);
  const put = (row: number, column: number, text: string) => {
    sheet.setValue(row, column, CellValue.text(text));
    (sheet.getOrCreate(row, column) as { styleId: number }).styleId = id;
  };
  return { wb, sheet, put };
}

describe('textHeight', () => {
  it('fits a default line in the default row', () => {
    expect(textHeight(Defaults.fontSizePoints)).toBeLessThanOrEqual(Defaults.rowHeight);
    expect(textHeight(20)).toBeGreaterThan(Defaults.rowHeight);
    expect(textHeight(10, 3)).toBeGreaterThan(textHeight(10, 2));
  });
});

describe('contentRowHeights', () => {
  it('leaves default rows out', () => {
    const wb = Workbook.createNew();
    wb.worksheets[0]!.setValue(0, 0, CellValue.text('hello'));
    expect(contentRowHeights(wb.worksheets[0]!, wb.styles, measure).size).toBe(0);
  });

  it('grows a row for a big font and for wrapped text', () => {
    const big = styled((s) => ({ ...s, font: { ...s.font, sizePoints: 24 } }));
    big.put(2, 0, 'Title');
    expect(contentRowHeights(big.sheet, big.wb.styles, measure).get(2)).toBe(textHeight(24));

    const wrap = styled((s) => ({ ...s, alignment: { ...s.alignment, wrapText: true } }));
    wrap.put(0, 0, 'x'.repeat(30)); // 180px in a 54px-wide cell → 4 lines
    expect(contentRowHeights(wrap.sheet, wrap.wb.styles, measure).get(0)).toBe(textHeight(10, 4));
  });

  it('grows a row to fit rotated text, by how far it is turned', () => {
    const m: RowMeasure = { ...measure, textWidth: (text) => text.length * 6 };
    const turned = (degrees: number) => {
      const t = styled((s) => ({ ...s, alignment: { ...s.alignment, rotation: degrees } }));
      t.put(0, 0, 'x'.repeat(20)); // 120px wide
      return contentRowHeights(t.sheet, t.wb.styles, m).get(0) ?? Defaults.rowHeight;
    };
    expect(turned(90)).toBe(120 + 2 * 2);
    expect(turned(270)).toBe(120 + 2 * 2);
    expect(turned(45)).toBeGreaterThan(Defaults.rowHeight);
    expect(turned(45)).toBeLessThan(turned(90));
  });

  it('counts typed line breaks, and leaves merged cells alone', () => {
    const wb = Workbook.createNew();
    const sheet = wb.worksheets[0]!;
    sheet.setValue(0, 0, CellValue.text('a\nb\nc'));
    sheet.setValue(4, 0, CellValue.text('a\nb'));
    sheet.addMerge(new MergedRegion(CellRange.parse('A5:B6')));
    const heights = contentRowHeights(sheet, wb.styles, measure);
    expect(heights.get(0)).toBe(textHeight(10, 3));
    expect(heights.has(4)).toBe(false);
  });
});

describe('ContentHeights.update', () => {
  it('re-measures only the edited cells and matches a full pass', () => {
    const big = styled((s) => ({ ...s, font: { ...s.font, sizePoints: 24 } }));
    big.put(2, 0, 'Title');
    big.put(2, 3, 'Other');
    const heights = new ContentHeights(big.sheet, big.wb.styles, measure);
    const tall = heights.rows.get(2)!;
    expect(tall).toBeGreaterThan(Defaults.rowHeight);

    // Clearing one of two tall cells keeps the row tall; clearing both shrinks it.
    big.sheet.remove(2, 0);
    expect(heights.update([{ row: 2, column: 0 }])).toBe(false);
    expect(heights.rows.get(2)).toBe(tall);
    big.sheet.remove(2, 3);
    expect(heights.update([{ row: 2, column: 3 }])).toBe(true);
    expect(heights.rows.has(2)).toBe(false);

    // A new tall cell grows its row, and the result equals measuring from scratch.
    big.put(7, 1, 'New');
    expect(heights.update([{ row: 7, column: 1 }])).toBe(true);
    expect([...heights.rows]).toEqual([...contentRowHeights(big.sheet, big.wb.styles, measure)]);
  });

  it('grows a row when typed text gets a line break', () => {
    const wb = Workbook.createNew();
    const sheet = wb.worksheets[0]!;
    const heights = new ContentHeights(sheet, wb.styles, measure);
    sheet.setValue(4, 0, CellValue.text('one\ntwo\nthree'));
    expect(heights.update([{ row: 4, column: 0 }])).toBe(true);
    expect(heights.rows.get(4)).toBe(textHeight(Defaults.fontSizePoints, 3));
  });
});
