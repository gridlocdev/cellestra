import { PAPER_SIZES, PageSetups, type PageSetup } from '@cellestra/core';

/** The Page Setup dialog's fields, as typed. */
export interface PageSetupDraft {
  /** A `PAPER_SIZES` id, or 'custom' for the setup's own size. */
  readonly paper: string;
  readonly orientation: 'portrait' | 'landscape';
  readonly top: string;
  readonly bottom: string;
  readonly left: string;
  readonly right: string;
  readonly scaleKind: 'percent' | 'fit' | 'pages';
  readonly percent: string;
  readonly fitWidth: string;
  readonly fitHeight: string;
  readonly pages: string;
  readonly gridlines: boolean;
  readonly headers: boolean;
  readonly centerHorizontally: boolean;
  readonly centerVertically: boolean;
  readonly pageOrder: 'downThenOver' | 'overThenDown';
}

const num = (v: number) => String(Math.round(v * 10) / 10);

export function draftOf(setup: PageSetup | null): PageSetupDraft {
  const s = setup ?? PageSetups.default;
  const scale = s.scale;
  return {
    paper: PageSetups.paperOf(s)?.id ?? 'custom',
    orientation: s.orientation,
    top: num(s.margins.top),
    bottom: num(s.margins.bottom),
    left: num(s.margins.left),
    right: num(s.margins.right),
    scaleKind: scale.kind,
    percent: scale.kind === 'percent' ? num(scale.percent) : '100',
    fitWidth: scale.kind === 'fit' ? (scale.width > 0 ? String(scale.width) : '') : '1',
    fitHeight: scale.kind === 'fit' && scale.height > 0 ? String(scale.height) : '',
    pages: scale.kind === 'pages' ? String(scale.pages) : '1',
    gridlines: s.gridlines,
    headers: s.headers,
    centerHorizontally: s.centerHorizontally,
    centerVertically: s.centerVertically,
    pageOrder: s.pageOrder,
  };
}

/** The setup the fields describe, over `base` (which gives a custom paper its size), or a problem. */
export function setupOf(
  draft: PageSetupDraft,
  base: PageSetup | null,
): { setup: PageSetup } | { error: string } {
  const from = base ?? PageSetups.default;
  const paper = PAPER_SIZES.find((p) => p.id === draft.paper);
  const sized = paper ? { ...from, paperWidth: paper.width, paperHeight: paper.height } : from;
  const oriented = PageSetups.orient(sized, draft.orientation);

  const margin = (text: string, label: string): number | string => {
    const v = Number(text.trim());
    return text.trim() !== '' && Number.isFinite(v) && v >= 0 && v <= 200
      ? v
      : `${label} margin must be 0–200 mm.`;
  };
  const margins = [
    margin(draft.top, 'The top'),
    margin(draft.bottom, 'The bottom'),
    margin(draft.left, 'The left'),
    margin(draft.right, 'The right'),
  ];
  const bad = margins.find((m) => typeof m === 'string');
  if (bad !== undefined) return { error: bad as string };
  const [top, bottom, left, right] = margins as number[];
  if (oriented.paperWidth - left! - right! < 20 || oriented.paperHeight - top! - bottom! < 20)
    return { error: 'The margins leave too little of the page to print on.' };

  const count = (text: string, allowBlank: boolean): number | null => {
    const t = text.trim();
    if (t === '') return allowBlank ? 0 : null;
    const v = Number(t);
    return Number.isInteger(v) && v >= (allowBlank ? 0 : 1) && v <= 999 ? v : null;
  };
  let scale: PageSetup['scale'];
  if (draft.scaleKind === 'percent') {
    const p = Number(draft.percent.trim());
    if (!Number.isFinite(p) || p < 10 || p > 400) return { error: 'The scale must be 10–400%.' };
    scale = { kind: 'percent', percent: p };
  } else if (draft.scaleKind === 'pages') {
    const pages = count(draft.pages, false);
    if (pages === null) return { error: 'Type how many pages to fit on (1 or more).' };
    scale = { kind: 'pages', pages };
  } else {
    const width = count(draft.fitWidth, true);
    const height = count(draft.fitHeight, true);
    if (width === null || height === null)
      return { error: 'Pages wide and tall are whole numbers (blank: any).' };
    if (width === 0 && height === 0) return { error: 'Give pages wide, pages tall, or both.' };
    scale = { kind: 'fit', width, height };
  }
  return {
    setup: {
      ...oriented,
      margins: { top: top!, bottom: bottom!, left: left!, right: right! },
      scale,
      gridlines: draft.gridlines,
      headers: draft.headers,
      centerHorizontally: draft.centerHorizontally,
      centerVertically: draft.centerVertically,
      pageOrder: draft.pageOrder,
    },
  };
}
