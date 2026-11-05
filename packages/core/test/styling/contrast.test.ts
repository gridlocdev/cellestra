import { describe, expect, it } from 'vitest';
import { Argb } from '../../src/styling/color.ts';
import { Contrast } from '../../src/styling/contrast.ts';

describe('Contrast', () => {
  it('black on white is the maximum and a colour on itself is the minimum', () => {
    expect(Contrast.ratio(Argb.black, Argb.white)).toBeCloseTo(21.0, 2);
    expect(Contrast.ratio(Argb.white, Argb.white)).toBeCloseTo(1.0, 6);
  });

  it('the ratio does not depend on which colour is the background', () => {
    const ink = Argb.fromRgb(0x15, 0x73, 0x4f);
    const paper = Argb.fromRgb(0xd9, 0xe5, 0xf5);
    expect(Contrast.ratio(ink, paper)).toBeCloseTo(Contrast.ratio(paper, ink), 9);
  });

  it('white on the table header band is below the readable floor', () => {
    // A header styled white over the light-theme header band is not low-contrast — it is invisible.
    expect(Contrast.ratio(Argb.white, Argb.fromRgb(0xd9, 0xe5, 0xf5))).toBeLessThan(3.0);
    expect(Contrast.ratio(Argb.black, Argb.fromRgb(0xd9, 0xe5, 0xf5))).toBeGreaterThan(3.0);
  });

  it.each([
    [0x00, 0.0],
    [0xff, 1.0],
  ])('luminance of channel %d is %d', (channel, expected) => {
    expect(Contrast.relativeLuminance(Argb.fromRgb(channel, channel, channel))).toBeCloseTo(expected, 6);
  });
});
