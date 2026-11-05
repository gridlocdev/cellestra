import { Argb, type ArgbColor } from './color.ts';

/** WCAG 2 contrast maths. */
export const Contrast = {
  ratio(a: ArgbColor, b: ArgbColor): number {
    const la = Contrast.relativeLuminance(a);
    const lb = Contrast.relativeLuminance(b);
    return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
  },
  relativeLuminance(c: ArgbColor): number {
    return 0.2126 * linear(Argb.r(c)) + 0.7152 * linear(Argb.g(c)) + 0.0722 * linear(Argb.b(c));
  },
} as const;

function linear(channel: number): number {
  const c = channel / 255;
  return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
}
