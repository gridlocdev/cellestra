import { CellStyles, Defaults, type CellStyle } from '@cellestra/core';
import { describe, expect, it } from 'vitest';
import { CELL_PAD_X, cellFont, fittedWidth, textHeight, type TextMeasure } from '../src/grid/cellText.ts';

/** A fake canvas: every character is 10px wide, and it remembers the font it was last given. */
function measure(): TextMeasure & { fonts: string[] } {
  const fonts: string[] = [];
  return {
    fonts,
    get font() {
      return fonts.at(-1) ?? '';
    },
    set font(f: string) {
      fonts.push(f);
    },
    measureText: (t) => ({ width: t.length * 10 }),
  };
}

const base = CellStyles.default;
const styled = (
  alignment: Partial<CellStyle['alignment']>,
  font: Partial<CellStyle['font']> = {},
): CellStyle => ({
  ...base,
  font: { ...base.font, ...font },
  alignment: { ...base.alignment, ...alignment },
});

describe('cell text', () => {
  it("fits flat text to its widest line plus the grid's padding each side", () => {
    expect(fittedWidth(measure(), 'abc\nabcdef', base)).toBe(60 + 2 * CELL_PAD_X);
  });

  it('measures in the font the grid draws with', () => {
    const ctx = measure();
    const bold = styled({}, { bold: true, family: 'Liberation Sans' });
    fittedWidth(ctx, 'x', bold);
    expect(ctx.fonts).toEqual([cellFont(bold.font)]);
    expect(cellFont(bold.font)).toMatch(/^bold \d+(\.\d+)?px "Liberation Sans", /);
    expect(cellFont(bold.font, 2)).toContain(`${bold.font.sizePoints * (96 / 72) * 2}px`);
  });

  it('asks no width for wrapped flat text, but measures turned text even when it wraps', () => {
    expect(fittedWidth(measure(), 'a long heading', styled({ wrapText: true }))).toBe(0);
    const upright = fittedWidth(measure(), 'a long heading', styled({ wrapText: true, rotation: 90 }));
    expect(upright).toBeGreaterThan(0);
    expect(upright).toBeLessThan(40);
  });

  it('makes a default-font line exactly the default row', () => {
    expect(textHeight(Defaults.fontSizePoints)).toBeLessThanOrEqual(Defaults.rowHeight);
  });
});
