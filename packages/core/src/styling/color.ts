/** A 32-bit ARGB colour stored as an unsigned number (0xAARRGGBB). */
export type ArgbColor = number;

export const Argb = {
  black: 0xff000000 as ArgbColor,
  white: 0xffffffff as ArgbColor,
  transparent: 0x00000000 as ArgbColor,

  a: (c: ArgbColor): number => (c >>> 24) & 0xff,
  r: (c: ArgbColor): number => (c >>> 16) & 0xff,
  g: (c: ArgbColor): number => (c >>> 8) & 0xff,
  b: (c: ArgbColor): number => c & 0xff,
  isTransparent: (c: ArgbColor): boolean => ((c >>> 24) & 0xff) === 0,

  fromRgb: (r: number, g: number, b: number): ArgbColor => ((0xff << 24) | (r << 16) | (g << 8) | b) >>> 0,
  fromArgb: (a: number, r: number, g: number, b: number): ArgbColor => ((a << 24) | (r << 16) | (g << 8) | b) >>> 0,

  /** `#RRGGBB` (alpha dropped), upper-case. */
  toHex(c: ArgbColor): string {
    return '#' + (c & 0xffffff).toString(16).toUpperCase().padStart(6, '0');
  },

  /** Parse `#RRGGBB` / `RRGGBB` as opaque; null when malformed. */
  tryFromHex(hex: string | null | undefined): ArgbColor | null {
    if (!hex) return null;
    let s = hex.trim();
    if (s.startsWith('#')) s = s.slice(1);
    if (!/^[0-9a-fA-F]{6}$/.test(s)) return null;
    return (0xff000000 | Number.parseInt(s, 16)) >>> 0;
  },

  fromHex(hex: string): ArgbColor {
    const c = Argb.tryFromHex(hex);
    if (c === null) throw new SyntaxError(`'${hex}' is not a #RRGGBB color.`);
    return c;
  },

  /** CSS `rgba()` for drawing. */
  toCss(c: ArgbColor): string {
    const a = Argb.a(c);
    return a === 255 ? Argb.toHex(c) : `rgba(${Argb.r(c)},${Argb.g(c)},${Argb.b(c)},${(a / 255).toFixed(3)})`;
  },
} as const;
