import { describe, expect, it } from 'vitest';
import { CellRange, CellValue, Defaults, MergedRegion, Workbook, type CellStyle } from '@cellestra/core';
import { contentRowHeights, textHeight, type RowMeasure } from '../src/grid/rowHeights.ts';

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
