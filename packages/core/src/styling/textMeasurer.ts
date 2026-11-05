import type { Font } from './cellStyle.ts';

/** Measures rendered text width in CSS px; the UI supplies a canvas-backed one for auto-fit. */
export interface TextMeasurer {
  measure(text: string, font: Font): number;
}

/** A per-character advance estimate, for headless use (tests, CLI). */
export const approximateTextMeasurer: TextMeasurer = {
  measure(text, font) {
    if (!text) return 0;
    const em = (font.sizePoints * 96) / 72;
    let units = 0;
    for (const c of text) units += advance(c);
    return units * em * (font.bold ? 1.04 : 1);
  },
};

function advance(c: string): number {
  if ("ijlI.,:;'`|!()[]/\\".includes(c)) return 0.28;
  if ('ftr -'.includes(c)) return 0.34;
  if ('mwMW@%'.includes(c)) return 0.88;
  if (c >= 'A' && c <= 'Z') return 0.68;
  if (c >= '0' && c <= '9') return 0.556;
  return 0.53;
}
