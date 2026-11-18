import { Argb, type ArgbColor } from '../../styling/color.ts';
import type { XmlElement } from '../../xml/index.ts';
import { at, kid } from './xlsxPackage.ts';

/** The legacy 64-colour palette `indexed="n"` points into (ECMA-376 §18.8.27). */
const INDEXED: readonly number[] = [
  0x000000, 0xffffff, 0xff0000, 0x00ff00, 0x0000ff, 0xffff00, 0xff00ff, 0x00ffff, 0x000000, 0xffffff,
  0xff0000, 0x00ff00, 0x0000ff, 0xffff00, 0xff00ff, 0x00ffff, 0x800000, 0x008000, 0x000080, 0x808000,
  0x800080, 0x008080, 0xc0c0c0, 0x808080, 0x9999ff, 0x993366, 0xffffcc, 0xccffff, 0x660066, 0xff8080,
  0x0066cc, 0xccccff, 0x000080, 0xff00ff, 0xffff00, 0x00ffff, 0x800080, 0x800000, 0x008080, 0x0000ff,
  0x00ccff, 0xccffff, 0xccffcc, 0xffff99, 0x99ccff, 0xff99cc, 0xcc99ff, 0xffcc99, 0x3366ff, 0x33cccc,
  0x99cc00, 0xffcc00, 0xff9900, 0xff6600, 0x666699, 0x969696, 0x003366, 0x339966, 0x003300, 0x333300,
  0x993300, 0x993366, 0x333399, 0x333333,
  // 64 and 65: the system foreground and background.
  0x000000, 0xffffff,
];

/** Office's default theme colours, in the order `theme="n"` counts them (lt1, dk1, lt2, dk2, accents, links). */
const OFFICE_THEME: readonly number[] = [
  0xffffff, 0x000000, 0xe7e6e6, 0x44546a, 0x4472c4, 0xed7d31, 0xa5a5a5, 0xffc000, 0x5b9bd5, 0x70ad47,
  0x0563c1, 0x954f72,
];

/** The scheme's element names in file order; `theme="n"` swaps the first two pairs (light first). */
const SCHEME_ORDER = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
];
const THEME_INDEX_TO_SCHEME = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11];

/** The colours a workbook's styles can name: its theme and its (possibly overridden) indexed palette. */
export interface XlsxPalette {
  readonly theme: readonly number[];
  readonly indexed: readonly number[];
}

/**
 * Read the palette from `theme1.xml` (`a:clrScheme`) and `styles.xml` (`colors/indexedColors`). A
 * missing theme falls back to Office's.
 */
export function readPalette(theme: XmlElement | null, styles: XmlElement | null): XlsxPalette {
  let themeColors = OFFICE_THEME;
  const scheme = theme?.descendants().find((e) => e.local === 'clrScheme') ?? null;
  if (scheme) {
    const byName = SCHEME_ORDER.map((name) => schemeColor(kid(scheme, name)));
    themeColors = THEME_INDEX_TO_SCHEME.map((i, n) => byName[i] ?? OFFICE_THEME[n]!);
  }
  let indexed = INDEXED;
  const custom = kid(kid(styles, 'colors'), 'indexedColors');
  if (custom) {
    const list = custom.children().filter((c) => c.local === 'rgbColor');
    indexed = INDEXED.map((c, i) => hexRgb(at(list[i] ?? null, 'rgb')) ?? c);
  }
  return { theme: themeColors, indexed };
}

function schemeColor(el: XmlElement | null): number | null {
  const color = el?.children()[0] ?? null;
  if (!color) return null;
  if (color.local === 'srgbClr') return hexRgb(at(color, 'val'));
  if (color.local === 'sysClr') return hexRgb(at(color, 'lastClr'));
  return null;
}

/** `RRGGBB` or `AARRGGBB` → 0xRRGGBB (alpha dropped: Excel ignores it), or null. */
function hexRgb(hex: string | null): number | null {
  if (hex === null || !/^([0-9a-f]{2})?[0-9a-f]{6}$/i.test(hex)) return null;
  return Number.parseInt(hex.slice(-6), 16);
}

/**
 * A `<color>`-shaped element (`rgb`, `theme`, `indexed`, `auto`, with an optional `tint`) as an
 * opaque colour, or null for "automatic" or nothing usable.
 */
export function readColor(el: XmlElement | null, palette: XlsxPalette): ArgbColor | null {
  if (!el || at(el, 'auto') === '1' || at(el, 'auto') === 'true') return null;
  let rgb: number | null = null;
  const theme = at(el, 'theme');
  const indexed = at(el, 'indexed');
  if (at(el, 'rgb') !== null) rgb = hexRgb(at(el, 'rgb'));
  else if (theme !== null) rgb = palette.theme[Number(theme)] ?? null;
  else if (indexed !== null) rgb = palette.indexed[Number(indexed)] ?? null;
  if (rgb === null) return null;
  const tint = Number(at(el, 'tint') ?? 0);
  return applyTint(rgb, Number.isFinite(tint) ? tint : 0);
}

/** Excel's tint: lighten (tint > 0) or darken (tint < 0) the colour's HSL lightness. */
export function applyTint(rgb: number, tint: number): ArgbColor {
  const r = (rgb >> 16) & 0xff;
  const g = (rgb >> 8) & 0xff;
  const b = rgb & 0xff;
  if (tint === 0) return Argb.fromRgb(r, g, b);
  const [h, s, l] = toHsl(r / 255, g / 255, b / 255);
  const lit = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const [r2, g2, b2] = fromHsl(h, s, Math.min(1, Math.max(0, lit)));
  return Argb.fromRgb(Math.round(r2 * 255), Math.round(g2 * 255), Math.round(b2 * 255));
}

function toHsl(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}

function fromHsl(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hue = (t: number): number => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [hue(h + 1 / 3), hue(h), hue(h - 1 / 3)];
}
