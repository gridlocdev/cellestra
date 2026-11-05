import type { ChartType } from '@cellestra/charts';
import type { ResolvedChart, ResolvedSeries } from '@cellestra/core';

/** An axis-aligned box in SVG user units (pixels). */
export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export const rectRight = (r: Rect): number => r.x + r.width;
export const rectBottom = (r: Rect): number => r.y + r.height;

/** A categorical palette with reasonable contrast in both themes; series cycle through it. */
export const SERIES_PALETTE: readonly string[] = [
  '#2b6cb0', // blue
  '#15734f', // green
  '#c25a1a', // orange
  '#8e44ad', // purple
  '#b02a37', // red
  '#1f9ea8', // teal
  '#b88a00', // gold
  '#556b2f', // olive
];

export function seriesColor(index: number): string {
  const n = SERIES_PALETTE.length;
  return SERIES_PALETTE[((index % n) + n) % n]!;
}

export const isBarType = (t: ChartType): boolean => t === 'column' || t === 'bar';

/** The value at `category`, or NaN past the end of the series (a gap). */
export function valueAt(series: ResolvedSeries, category: number): number {
  return category < series.values.length ? series.values[category]! : NaN;
}

/** Sum of the positive, non-gap values at `category` across `series`. */
export function positiveSum(series: readonly ResolvedSeries[], category: number): number {
  let sum = 0;
  for (const s of series) {
    const v = valueAt(s, category);
    if (!Number.isNaN(v) && v > 0) sum += v;
  }
  return sum;
}

/** The longest series' length: how many categories a category-axis chart spans. */
export function categoryCount(chart: ResolvedChart): number {
  let n = 0;
  for (const s of chart.series) n = Math.max(n, s.values.length);
  return n;
}

/** The label for category `c`: the resolved category text, else its 1-based ordinal. */
export function categoryLabel(chart: ResolvedChart, c: number): string {
  return c < chart.categories.length ? chart.categories[c]! : String(c + 1);
}

/**
 * The value-axis extent for a category chart, always including zero. Stacked bars contribute their
 * positive and negative running sums, stacked areas their positive running sum, everything else
 * its raw values. 100%-stacked charts span 0..1.
 */
export function valueExtent(
  chart: ResolvedChart,
  categories: number,
  stacked: boolean,
  percent: boolean,
): { min: number; max: number } {
  if (percent) return { min: 0, max: 1 };
  let min = 0;
  let max = 0;
  const bars = chart.series.filter((s) => isBarType(s.type));
  const areas = chart.series.filter((s) => s.type === 'area');
  const others = chart.series.filter((s) => !isBarType(s.type) && !(stacked && s.type === 'area'));

  for (let c = 0; c < categories; c++) {
    if (stacked && bars.length > 0) {
      let pos = 0;
      let neg = 0;
      for (const s of bars) {
        const v = valueAt(s, c);
        if (Number.isNaN(v)) continue;
        if (v >= 0) pos += v;
        else neg += v;
      }
      max = Math.max(max, pos);
      min = Math.min(min, neg);
    } else {
      for (const s of bars) {
        const v = valueAt(s, c);
        if (Number.isNaN(v)) continue;
        max = Math.max(max, v);
        min = Math.min(min, v);
      }
    }
    if (stacked && areas.length > 0) max = Math.max(max, positiveSum(areas, c));
    for (const s of others) {
      const v = valueAt(s, c);
      if (Number.isNaN(v)) continue;
      max = Math.max(max, v);
      min = Math.min(min, v);
    }
  }
  if (max <= min) max = min + 1;
  return { min, max };
}

/** Evenly spaced axis ticks on "nice" round numbers (1, 2, 2.5, 5 × 10ⁿ) that cover the extent. */
export interface Ticks {
  readonly min: number;
  readonly max: number;
  readonly step: number;
  readonly values: readonly number[];
}

/** Drop binary floating-point noise (0.30000000000000004 → 0.3). */
const clean = (n: number): number => (n === 0 ? 0 : Number(n.toPrecision(12)));

/** The nice step for a raw interval: the smallest of 1, 2, 2.5, 5, 10 × 10ⁿ that is ≥ it. */
export function niceStep(raw: number): number {
  if (!(raw > 0) || !Number.isFinite(raw)) return 1;
  const magnitude = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / magnitude;
  const nice = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10;
  return clean(nice * magnitude);
}

/** About `target` intervals between nice bounds that contain [min, max]. */
export function niceTicks(min: number, max: number, target = 5): Ticks {
  if (!Number.isFinite(min) || !Number.isFinite(max)) return niceTicks(0, 1, target);
  if (max < min) [min, max] = [max, min];
  if (Math.abs(max - min) < 1e-12) max = min + 1;
  const span = max - min;
  // Spans past the double range (±1e308) or below its precision at this magnitude can't be split
  // into steps: fall back to the bare extent instead of NaN ticks or an unbounded loop.
  if (!Number.isFinite(span) || span <= Math.max(Math.abs(min), Math.abs(max)) * 1e-12)
    return { min, max, step: Number.isFinite(span) ? span : max, values: [min, max] };
  const step = niceStep(span / Math.max(1, target));
  const lo = clean(Math.floor(clean(min / step)) * step);
  const hi = clean(Math.ceil(clean(max / step)) * step);
  const count = Math.round((hi - lo) / step);
  if (!Number.isFinite(lo) || !Number.isFinite(hi) || !(count >= 1) || count > 1000)
    return { min, max, step: span, values: [min, max] };
  const values: number[] = [];
  for (let i = 0; i <= count; i++) values.push(clean(lo + i * step));
  return { min: lo, max: hi, step, values };
}

/** Pixel position of `value` along the value axis of `area` (bottom-up, or left-to-right). */
export function valuePos(value: number, min: number, max: number, area: Rect, horizontal: boolean): number {
  const t = (value - min) / (max - min);
  return horizontal ? area.x + t * area.width : rectBottom(area) - t * area.height;
}

/** .NET `"#,##0"` for |v| ≥ 1000, else `"0.##"` (invariant culture, half away from zero). */
export function formatNumber(value: number): string {
  if (!Number.isFinite(value)) return String(value);
  const abs = Math.abs(value);
  if (abs >= 1000) {
    const whole = Math.round(abs);
    const body = whole.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return value < 0 ? '-' + body : body;
  }
  const rounded = Math.round(abs * 100) / 100;
  if (rounded === 0) return '0';
  const body = String(clean(rounded));
  return value < 0 ? '-' + body : body;
}

/** A value-axis tick label: `20%` on a 100%-stacked axis, else {@link formatNumber}. */
export function formatTick(value: number, percent: boolean): string {
  return percent ? `${Math.round(value * 100)}%` : formatNumber(value);
}

/**
 * Rough rendered width of `text` in the UI font. SVG text can't be measured without a DOM, so
 * layout uses an average glyph width; it only has to be close enough to reserve gutters.
 */
export function estimateTextWidth(text: string, size: number, bold = false): number {
  return text.length * size * (bold ? 0.6 : 0.56);
}

/** One stacked bar segment, in axis units. */
export interface StackSegment {
  readonly seriesIndex: number;
  readonly category: number;
  readonly start: number;
  readonly end: number;
}

/**
 * Stack the values of `series` per category, in series order: positives upward from zero, negatives
 * downward (gaps skipped). With `percent` each category is scaled so its positives sum to 1 and
 * negatives are dropped.
 * `indices[i]` is `series[i]`'s index in the chart (its colour).
 */
export function stackBars(
  series: readonly ResolvedSeries[],
  indices: readonly number[],
  categories: number,
  percent: boolean,
): StackSegment[] {
  const out: StackSegment[] = [];
  for (let c = 0; c < categories; c++) {
    const total = percent ? positiveSum(series, c) : 1;
    let cursor = 0;
    let below = 0;
    series.forEach((s, i) => {
      const v = valueAt(s, c);
      if (Number.isNaN(v) || v === 0) return;
      if (v < 0) {
        // Negatives stack downward from zero (the axis extent already makes room for them);
        // 100% stacking scales positives only.
        if (percent) return;
        out.push({ seriesIndex: indices[i] ?? i, category: c, start: below, end: below + v });
        below += v;
        return;
      }
      const display = percent ? (total > 0 ? v / total : 0) : v;
      const start = percent ? (total > 0 ? cursor / total : 0) : cursor;
      out.push({ seriesIndex: indices[i] ?? i, category: c, start, end: start + display });
      cursor += v;
    });
  }
  return out;
}

/**
 * The base and top of each area series per category. Unstacked areas rise from zero to their value
 * (gaps as zero); stacked areas sit on the running sum of the positive values beneath them, scaled
 * to 0..1 per category when `percent`.
 */
export function stackAreas(
  series: readonly ResolvedSeries[],
  categories: number,
  stacked: boolean,
  percent: boolean,
): { base: number[]; top: number[] }[] {
  const out = series.map(() => ({ base: [] as number[], top: [] as number[] }));
  for (let c = 0; c < categories; c++) {
    const total = percent ? positiveSum(series, c) : 1;
    let cursor = 0;
    series.forEach((s, i) => {
      const raw = valueAt(s, c);
      const v = Number.isNaN(raw) ? 0 : raw;
      const lane = out[i]!;
      if (!stacked) {
        lane.base.push(0);
        lane.top.push(v);
        return;
      }
      const add = Math.max(0, v);
      const scale = percent ? (total > 0 ? 1 / total : 0) : 1;
      lane.base.push(cursor * scale);
      lane.top.push((cursor + add) * scale);
      cursor += add;
    });
  }
  return out;
}

/** One pie/doughnut wedge: angles in radians, clockwise from 12 o'clock (−π/2). */
export interface PieSlice {
  readonly index: number;
  readonly value: number;
  readonly start: number;
  readonly end: number;
}

/** The first series' positive values as wedges; empty when nothing is positive. */
export function pieSlices(chart: ResolvedChart): PieSlice[] {
  const series = chart.series[0];
  if (!series) return [];
  const raw: { index: number; value: number }[] = [];
  series.values.forEach((v, index) => {
    if (!Number.isNaN(v) && v > 0) raw.push({ index, value: v });
  });
  const total = raw.reduce((sum, s) => sum + s.value, 0);
  if (total <= 0) return [];
  let angle = -Math.PI / 2;
  return raw.map((s) => {
    const sweep = (s.value / total) * Math.PI * 2;
    const slice = { index: s.index, value: s.value, start: angle, end: angle + sweep };
    angle += sweep;
    return slice;
  });
}

const fmt = (n: number): string => String(Math.round(n * 100) / 100);

/** SVG path data for a wedge (or a doughnut ring segment when `innerRadius` > 0). */
export function wedgePath(
  cx: number,
  cy: number,
  radius: number,
  innerRadius: number,
  start: number,
  end: number,
): string {
  // A full circle can't be one arc (start == end point), so split it in two halves.
  if (end - start >= Math.PI * 2 - 1e-9) {
    const mid = start + Math.PI;
    return (
      wedgePath(cx, cy, radius, innerRadius, start, mid) + wedgePath(cx, cy, radius, innerRadius, mid, end)
    );
  }
  const large = end - start > Math.PI ? 1 : 0;
  const pt = (r: number, a: number) => `${fmt(cx + r * Math.cos(a))} ${fmt(cy + r * Math.sin(a))}`;
  if (innerRadius <= 0) {
    return `M${fmt(cx)} ${fmt(cy)}L${pt(radius, start)}A${fmt(radius)} ${fmt(radius)} 0 ${large} 1 ${pt(radius, end)}Z`;
  }
  return (
    `M${pt(innerRadius, start)}L${pt(radius, start)}` +
    `A${fmt(radius)} ${fmt(radius)} 0 ${large} 1 ${pt(radius, end)}` +
    `L${pt(innerRadius, end)}` +
    `A${fmt(innerRadius)} ${fmt(innerRadius)} 0 ${large} 0 ${pt(innerRadius, start)}Z`
  );
}

/** A legend entry: label plus the palette index of its swatch. */
export interface LegendItem {
  readonly name: string;
  readonly index: number;
}

/**
 * Pie/doughnut legends name the drawn slices (categories), keyed by slice index; every other type
 * names its series.
 */
export function legendItems(chart: ResolvedChart): LegendItem[] {
  if (chart.type === 'pie' || chart.type === 'doughnut') {
    return pieSlices(chart).map((s) => ({ name: categoryLabel(chart, s.index), index: s.index }));
  }
  return chart.series.map((s, index) => ({ name: s.name, index }));
}

export const LEGEND_ROW_HEIGHT = 26;
const LEGEND_SPACING = 18;
const LEGEND_SWATCH = 14 + 4; // swatch + gap before the label

/** Legend entries placed in centred rows (wrapping when they don't fit `width`). */
export interface LegendLayout {
  readonly height: number;
  readonly items: readonly { item: LegendItem; x: number; row: number }[];
}

export function layoutLegend(items: readonly LegendItem[], width: number, fontSize = 11): LegendLayout {
  if (items.length === 0) return { height: 0, items: [] };
  const rows: { item: LegendItem; w: number }[][] = [[]];
  let used = 0;
  for (const item of items) {
    const w = LEGEND_SWATCH + estimateTextWidth(item.name, fontSize);
    let row = rows[rows.length - 1]!;
    const needed = row.length === 0 ? w : used + LEGEND_SPACING + w;
    if (row.length > 0 && needed > width) {
      row = [];
      rows.push(row);
      used = w;
    } else {
      used = needed;
    }
    row.push({ item, w });
  }
  const placed: { item: LegendItem; x: number; row: number }[] = [];
  rows.forEach((row, r) => {
    const total = row.reduce((sum, e) => sum + e.w, 0) + LEGEND_SPACING * (row.length - 1);
    let x = Math.max(0, (width - total) / 2);
    for (const e of row) {
      placed.push({ item: e.item, x, row: r });
      x += e.w + LEGEND_SPACING;
    }
  });
  return { height: rows.length * LEGEND_ROW_HEIGHT, items: placed };
}
