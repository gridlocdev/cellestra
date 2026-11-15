import type { ChartAggregate } from '@cellestra/charts';
import type { ResolvedChart, ResolvedSeries } from '@cellestra/core';
import { categoryLabel, formatNumber, isBarType, valueAt, type Rect } from './chartMath.ts';
import type { Shape } from './chartScene.ts';

/** One line of a tooltip: a label and its formatted value, with an optional series swatch. */
export interface HoverRow {
  readonly label: string;
  readonly value: string;
  /** Palette index for a colour swatch. */
  readonly series?: number;
  /** The row for the mark under the pointer. */
  readonly emphasis?: boolean;
}

/** What a tooltip says about the hovered part of a chart. */
export interface HoverInfo {
  readonly title: string;
  readonly rows: readonly HoverRow[];
  /** A footnote, such as "Sum of 3 rows" for a category that merged rows. */
  readonly note?: string;
}

export interface Point {
  readonly x: number;
  readonly y: number;
}

/** Where a hit answers the pointer, in chart coordinates. Angles as in `pieSlices`. */
export type HitRegion =
  | {
      readonly kind: 'rect';
      readonly x: number;
      readonly y: number;
      readonly width: number;
      readonly height: number;
    }
  | {
      readonly kind: 'wedge';
      readonly cx: number;
      readonly cy: number;
      readonly inner: number;
      readonly outer: number;
      readonly start: number;
      readonly end: number;
    }
  /** The nearest point within `radius` wins. */
  | { readonly kind: 'point'; readonly x: number; readonly y: number; readonly radius: number };

/** A bar inside a category band: the one nearest the pointer is emphasised. */
export interface HoverBar {
  readonly rect: Rect;
  readonly series: number;
  /** Where the tooltip points when this bar is emphasised. */
  readonly anchor: Point;
}

/** A hover target the scene lays out next to its shapes. */
export interface ChartHit {
  readonly region: HitRegion;
  readonly info: HoverInfo;
  readonly anchor: Point;
  /** Shapes drawn over the chart while the hit is hovered. */
  readonly highlight: readonly Shape[];
  /** Bars in a category band; `horizontal` says the band runs along y (a bar chart). */
  readonly bars?: readonly HoverBar[];
  readonly horizontal?: boolean;
}

/** What to show for the pointer's position. */
export interface Hover {
  readonly info: HoverInfo;
  readonly anchor: Point;
  readonly highlight: readonly Shape[];
}

/** How close (px) the pointer must be to a scatter point to hover it. */
export const POINT_HIT_RADIUS = 12;
/** A gap (blank or non-numeric cell). */
export const GAP_TEXT = '—';
const TOOLTIP_GAP = 10;
const TOOLTIP_MARGIN = 4;
const TAU = Math.PI * 2;

const AGGREGATE_WORDS: Readonly<Record<ChartAggregate, string>> = {
  sum: 'Sum',
  average: 'Average',
  count: 'Count',
  min: 'Smallest',
  max: 'Largest',
};

export const formatValue = (v: number): string => (Number.isNaN(v) ? GAP_TEXT : formatNumber(v));

/** A share as a percentage with at most one decimal: 0.3333 → "33.3%". */
export const formatPercent = (share: number): string => `${Math.round(share * 1000) / 10}%`;

/** "Sum of 3 rows" when category `c` merged more than one row, else undefined. */
export function groupNote(chart: ResolvedChart, c: number): string | undefined {
  const rows = chart.rowCounts?.[c] ?? 1;
  if (rows <= 1) return undefined;
  return `${AGGREGATE_WORDS[chart.aggregate ?? 'sum']} of ${rows} rows`;
}

/** The family a series stacks with on a category chart (other types draw as the chart's type). */
function stackFamily(s: ResolvedSeries, chart: ResolvedChart): 'bar' | 'line' | 'area' {
  const type = isBarType(s.type) || s.type === 'line' || s.type === 'area' ? s.type : chart.type;
  return isBarType(type) ? 'bar' : type === 'line' ? 'line' : 'area';
}

/**
 * Category `c` of a column, bar, line, area or radar chart: every series' value there. Stacked
 * charts add the total; 100% stacked charts also give each value's share of its stack.
 */
export function categoryInfo(chart: ResolvedChart, c: number): HoverInfo {
  const stacked = chart.stacking === 'stacked' || chart.stacking === 'percent';
  const percent = chart.stacking === 'percent';
  const families = chart.series.map((s) => stackFamily(s, chart));
  const rows: HoverRow[] = chart.series.map((s, i) => {
    const v = valueAt(s, c);
    let value = formatValue(v);
    if (percent && v > 0) {
      let total = 0;
      chart.series.forEach((other, k) => {
        const o = valueAt(other, c);
        if (families[k] === families[i] && o > 0) total += o;
      });
      if (total > 0) value += ` (${formatPercent(v / total)})`;
    }
    return { label: s.name, value, series: i };
  });
  if (stacked && chart.series.length > 1) {
    let total = 0;
    for (const s of chart.series) {
      const v = valueAt(s, c);
      if (!Number.isNaN(v)) total += v;
    }
    rows.push({ label: 'Total', value: formatNumber(total) });
  }
  return { title: categoryLabel(chart, c), rows, note: groupNote(chart, c) };
}

/** A pie or doughnut slice: its value and share of the whole. */
export function sliceInfo(chart: ResolvedChart, index: number, value: number, total: number): HoverInfo {
  const name = chart.series[0]?.name ?? '';
  return {
    title: categoryLabel(chart, index),
    rows: [
      { label: name, value: formatNumber(value), series: index },
      { label: 'Share', value: formatPercent(total > 0 ? value / total : 0) },
    ],
    note: groupNote(chart, index),
  };
}

/** A scatter or bubble point: X and Y (named by the axis titles when set), and size for a bubble. */
export function pointInfo(
  chart: ResolvedChart,
  seriesIndex: number,
  x: number,
  y: number,
  size: number | null,
): HoverInfo {
  const rows: HoverRow[] = [
    { label: chart.xAxisTitle?.trim() || 'X', value: formatValue(x) },
    { label: chart.yAxisTitle?.trim() || 'Y', value: formatValue(y), series: seriesIndex },
  ];
  if (size !== null) rows.push({ label: 'Size', value: formatValue(size) });
  return { title: chart.series[seriesIndex]?.name ?? '', rows };
}

function wedgeContains(r: Extract<HitRegion, { kind: 'wedge' }>, x: number, y: number): boolean {
  const dx = x - r.cx;
  const dy = y - r.cy;
  const d = Math.hypot(dx, dy);
  if (d < r.inner || d > r.outer) return false;
  const rel = (((Math.atan2(dy, dx) - r.start) % TAU) + TAU) % TAU;
  return rel < r.end - r.start;
}

function regionContains(r: HitRegion, x: number, y: number): boolean {
  switch (r.kind) {
    case 'rect':
      return x >= r.x && x < r.x + r.width && y >= r.y && y < r.y + r.height;
    case 'wedge':
      return wedgeContains(r, x, y);
    case 'point':
      return Math.hypot(x - r.x, y - r.y) <= r.radius;
  }
}

/** Distance from (x, y) to `r` (0 inside). */
function rectDistance(r: Rect, x: number, y: number): number {
  const dx = Math.max(r.x - x, 0, x - (r.x + r.width));
  const dy = Math.max(r.y - y, 0, y - (r.y + r.height));
  return Math.hypot(dx, dy);
}

/**
 * The bar to emphasise: of the bars whose slot along the category axis holds the pointer, the
 * nearest; else the nearest of all (stacked bars share one slot, so the nearest segment wins).
 */
function pickBar(bars: readonly HoverBar[], horizontal: boolean, x: number, y: number): HoverBar | null {
  const along = (b: HoverBar) =>
    horizontal
      ? y >= b.rect.y && y <= b.rect.y + b.rect.height
      : x >= b.rect.x && x <= b.rect.x + b.rect.width;
  const inSlot = bars.filter(along);
  const pool = inSlot.length > 0 ? inSlot : bars;
  let best: HoverBar | null = null;
  let bestD = Infinity;
  for (const b of pool) {
    const d = rectDistance(b.rect, x, y);
    if (d < bestD) {
      best = b;
      bestD = d;
    }
  }
  return best;
}

const rectPath = (r: Rect): string => `M${r.x} ${r.y}h${r.width}v${r.height}h${-r.width}Z`;

function resolve(hit: ChartHit, x: number, y: number): Hover {
  const bar = hit.bars && hit.bars.length > 0 ? pickBar(hit.bars, hit.horizontal ?? false, x, y) : null;
  if (!bar) return { info: hit.info, anchor: hit.anchor, highlight: hit.highlight };
  return {
    info: {
      ...hit.info,
      rows: hit.info.rows.map((r) => (r.series === bar.series ? { ...r, emphasis: true } : r)),
    },
    anchor: bar.anchor,
    highlight: [
      ...hit.highlight,
      { kind: 'path', d: rectPath(bar.rect), fill: null, stroke: 'text', strokeWidth: 2 },
    ],
  };
}

/**
 * What the pointer at (x, y) is over: the first band, slice or spoke that holds it, else the
 * nearest point within its radius; null over nothing.
 */
export function hitTest(scene: { readonly hits: readonly ChartHit[] }, x: number, y: number): Hover | null {
  let nearest: ChartHit | null = null;
  let nearestD = Infinity;
  for (const hit of scene.hits) {
    const r = hit.region;
    if (r.kind === 'point') {
      const d = Math.hypot(x - r.x, y - r.y);
      if (d <= r.radius && d < nearestD) {
        nearest = hit;
        nearestD = d;
      }
      continue;
    }
    if (regionContains(r, x, y)) return resolve(hit, x, y);
  }
  return nearest ? resolve(nearest, x, y) : null;
}

/**
 * The tooltip's top-left for a `tip`-sized box pointing at `anchor` inside `box`: centred above the
 * anchor, below it when there is no room above, and kept inside the box.
 */
export function placeTooltip(
  anchor: Point,
  tip: { readonly width: number; readonly height: number },
  box: { readonly width: number; readonly height: number },
): Point {
  const clamp = (v: number, lo: number, hi: number) => Math.max(lo, Math.min(v, hi));
  const left = clamp(anchor.x - tip.width / 2, TOOLTIP_MARGIN, box.width - tip.width - TOOLTIP_MARGIN);
  let top = anchor.y - TOOLTIP_GAP - tip.height;
  if (top < TOOLTIP_MARGIN) top = anchor.y + TOOLTIP_GAP;
  top = clamp(top, TOOLTIP_MARGIN, box.height - tip.height - TOOLTIP_MARGIN);
  return { x: Math.max(TOOLTIP_MARGIN, left), y: Math.max(TOOLTIP_MARGIN, top) };
}
