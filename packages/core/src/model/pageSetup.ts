/** How a printed sheet is sized: a zoom, or shrunk to fit a number of pages. */
export type PageScale =
  | { readonly kind: 'percent'; readonly percent: number }
  /** At most `pages` pages in all (ODF `style:scale-to-pages`). */
  | { readonly kind: 'pages'; readonly pages: number }
  /** At most `width` pages across and `height` pages down; 0 leaves that direction free. */
  | { readonly kind: 'fit'; readonly width: number; readonly height: number };

export type PageOrientation = 'portrait' | 'landscape';

export interface PageMargins {
  readonly top: number;
  readonly bottom: number;
  readonly left: number;
  readonly right: number;
}

/**
 * How a sheet prints (ODF `style:page-layout`). Lengths are millimetres; the paper size is as the
 * page lies, so a landscape A4 page is 297 mm wide.
 */
export interface PageSetup {
  readonly paperWidth: number;
  readonly paperHeight: number;
  readonly orientation: PageOrientation;
  readonly margins: PageMargins;
  readonly scale: PageScale;
  /** Print the cell grid lines. */
  readonly gridlines: boolean;
  /** Print the row numbers and column letters. */
  readonly headers: boolean;
  readonly centerHorizontally: boolean;
  readonly centerVertically: boolean;
  /** Which way pages run when the sheet is wider and taller than one page. */
  readonly pageOrder: 'downThenOver' | 'overThenDown';
}

export interface PaperSize {
  readonly id: string;
  readonly label: string;
  /** Portrait width and height, mm. */
  readonly width: number;
  readonly height: number;
}

export const PAPER_SIZES: readonly PaperSize[] = [
  { id: 'a4', label: 'A4', width: 210, height: 297 },
  { id: 'letter', label: 'Letter', width: 215.9, height: 279.4 },
  { id: 'legal', label: 'Legal', width: 215.9, height: 355.6 },
  { id: 'a3', label: 'A3', width: 297, height: 420 },
  { id: 'a5', label: 'A5', width: 148, height: 210 },
  { id: 'tabloid', label: 'Tabloid', width: 279.4, height: 431.8 },
];

export const PageSetups = {
  /** LibreOffice's default: A4 portrait, 2 cm margins, 100%. */
  default: Object.freeze({
    paperWidth: 210,
    paperHeight: 297,
    orientation: 'portrait',
    margins: Object.freeze({ top: 20, bottom: 20, left: 20, right: 20 }),
    scale: Object.freeze({ kind: 'percent', percent: 100 }),
    gridlines: false,
    headers: false,
    centerHorizontally: false,
    centerVertically: false,
    pageOrder: 'downThenOver',
  }) as PageSetup,

  /** The paper size a setup uses, matched within 1 mm either way round; null when it is custom. */
  paperOf(setup: PageSetup): PaperSize | null {
    const near = (a: number, b: number) => Math.abs(a - b) < 1;
    return (
      PAPER_SIZES.find(
        (p) =>
          (near(p.width, setup.paperWidth) && near(p.height, setup.paperHeight)) ||
          (near(p.height, setup.paperWidth) && near(p.width, setup.paperHeight)),
      ) ?? null
    );
  },

  /** The setup with its paper turned to `orientation` (width and height swapped when needed). */
  orient(setup: PageSetup, orientation: PageOrientation): PageSetup {
    const short = Math.min(setup.paperWidth, setup.paperHeight);
    const long = Math.max(setup.paperWidth, setup.paperHeight);
    return orientation === 'portrait'
      ? { ...setup, orientation, paperWidth: short, paperHeight: long }
      : { ...setup, orientation, paperWidth: long, paperHeight: short };
  },

  equal(a: PageSetup | null, b: PageSetup | null): boolean {
    if (a === null || b === null) return a === b;
    const m = (x: PageMargins) => [x.top, x.bottom, x.left, x.right];
    const s = (x: PageScale) =>
      x.kind === 'percent'
        ? [x.kind, x.percent]
        : x.kind === 'pages'
          ? [x.kind, x.pages]
          : [x.kind, x.width, x.height];
    const flat = (x: PageSetup) =>
      JSON.stringify([
        x.paperWidth,
        x.paperHeight,
        x.orientation,
        m(x.margins),
        s(x.scale),
        x.gridlines,
        x.headers,
        x.centerHorizontally,
        x.centerVertically,
        x.pageOrder,
      ]);
    return flat(a) === flat(b);
  },
} as const;
