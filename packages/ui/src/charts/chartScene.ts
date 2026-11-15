import type { ResolvedChart, ResolvedSeries } from '@cellestra/core';
import {
  type Rect,
  categoryCount,
  categoryLabel,
  estimateTextWidth,
  formatNumber,
  formatTick,
  isBarType,
  layoutLegend,
  LEGEND_ROW_HEIGHT,
  legendItems,
  niceTicks,
  pieSlices,
  rectBottom,
  rectRight,
  stackAreas,
  stackBars,
  stackLines,
  valueAt,
  valueExtent,
  valuePos,
  wedgePath,
} from './chartMath.ts';
import {
  type ChartHit,
  type HoverBar,
  categoryInfo,
  POINT_HIT_RADIUS,
  pointInfo,
  sliceInfo,
} from './chartHover.ts';

/**
 * A theme role (resolved to a CSS variable by the stylesheet) or a series palette index. Keeping
 * colours symbolic lets the scene stay pure and the SVG follow the app theme.
 */
export type Paint = 'bg' | 'text' | 'muted' | 'axis' | 'grid' | { readonly series: number };

export type Shape =
  | {
      readonly kind: 'rect';
      readonly x: number;
      readonly y: number;
      readonly width: number;
      readonly height: number;
      readonly fill: Paint | null;
      readonly stroke?: Paint;
    }
  | {
      readonly kind: 'line';
      readonly x1: number;
      readonly y1: number;
      readonly x2: number;
      readonly y2: number;
      readonly stroke: Paint;
      readonly width: number;
    }
  | {
      readonly kind: 'path';
      readonly d: string;
      readonly fill: Paint | null;
      readonly fillOpacity?: number;
      readonly stroke: Paint | null;
      readonly strokeWidth?: number;
    }
  | {
      readonly kind: 'circle';
      readonly cx: number;
      readonly cy: number;
      readonly r: number;
      readonly fill: Paint | null;
      readonly opacity?: number;
      readonly stroke?: Paint;
      readonly strokeWidth?: number;
    }
  | {
      readonly kind: 'text';
      readonly x: number;
      readonly y: number;
      readonly text: string;
      readonly size: number;
      readonly paint: Paint;
      readonly anchor: 'start' | 'middle' | 'end';
      readonly bold?: boolean;
      /** Rotation in degrees about (x, y). */
      readonly rotate?: number;
    };

/** Everything the chart draws, in paint order, with `y` on text meaning its vertical centre. */
export interface ChartScene {
  readonly width: number;
  readonly height: number;
  readonly shapes: readonly Shape[];
  /** Hover targets for {@link hitTest}, in the same coordinates as the shapes. */
  readonly hits: readonly ChartHit[];
}

const r2 = (n: number): number => Math.round(n * 100) / 100;

const TITLE_SIZE = 16;
const AXIS_TITLE_SIZE = 11;
const AXIS_TITLE_GUTTER = 18;
const TICK_SIZE = 10;
const CATEGORY_SIZE = 11;

/** Lay out a resolved chart in a `width` × `height` box. Null `chart` shows "No chart". */
export function buildChartScene(chart: ResolvedChart | null, width: number, height: number): ChartScene {
  const shapes: Shape[] = [];
  const hits: ChartHit[] = [];
  const bounds: Rect = { x: 0, y: 0, width, height };
  shapes.push({ kind: 'rect', ...bounds, fill: 'bg' });

  if (!chart || chart.isEmpty) {
    centered(shapes, chart ? 'No data — set the ranges in the panel.' : 'No chart', bounds, 14);
    return { width, height, shapes, hits };
  }

  let top = 12;
  if (chart.title.trim() !== '') {
    const h = TITLE_SIZE * 1.3;
    shapes.push({
      kind: 'text',
      x: width / 2,
      y: top + h / 2,
      text: chart.title,
      size: TITLE_SIZE,
      paint: 'text',
      anchor: 'middle',
      bold: true,
    });
    top += h + 8;
  }

  const legend = chart.showLegend ? layoutLegend(legendItems(chart), width) : null;
  const bottomReserve = legend?.height ?? 0;
  const plot: Rect = { x: 0, y: top, width, height: Math.max(0, height - top - bottomReserve - 8) };

  if (plot.width >= 40 && plot.height >= 40) {
    switch (chart.type) {
      case 'pie':
      case 'doughnut':
        drawPie(shapes, hits, chart, plot, chart.type === 'doughnut');
        break;
      case 'scatter':
      case 'bubble':
        drawScatter(shapes, hits, chart, plot, chart.type === 'bubble');
        break;
      case 'radar':
      case 'filledRadar':
        drawRadar(shapes, hits, chart, plot, chart.type === 'filledRadar');
        break;
      default:
        drawCartesian(shapes, hits, chart, plot);
        break;
    }
  }

  if (legend && legend.items.length > 0) {
    const y0 = height - bottomReserve;
    for (const { item, x, row } of legend.items) {
      const cy = y0 + row * LEGEND_ROW_HEIGHT + LEGEND_ROW_HEIGHT / 2;
      shapes.push({ kind: 'rect', x, y: cy - 6, width: 12, height: 12, fill: { series: item.index } });
      shapes.push({
        kind: 'text',
        x: x + 16,
        y: cy,
        text: item.name,
        size: 11,
        paint: 'text',
        anchor: 'start',
      });
    }
  }

  return { width, height, shapes, hits };
}

function centered(shapes: Shape[], message: string, area: Rect, size: number): void {
  shapes.push({
    kind: 'text',
    x: area.x + area.width / 2,
    y: area.y + area.height / 2,
    text: message,
    size,
    paint: 'muted',
    anchor: 'middle',
  });
}

// ---- Cartesian (column, bar, line, area, combos) ---------------------------------------------

/** The inner plotting area once gutters for tick labels and axis titles are taken out of `plot`. */
function axisArea(chart: ResolvedChart, plot: Rect, left: number, bottom: number): Rect {
  const l = left + (chart.yAxisTitle?.trim() ? AXIS_TITLE_GUTTER : 0);
  const b = bottom + (chart.xAxisTitle?.trim() ? AXIS_TITLE_GUTTER : 0);
  return {
    x: plot.x + l,
    y: plot.y + 6,
    width: Math.max(0, plot.width - l - 12),
    height: Math.max(0, plot.height - b - 6),
  };
}

const tickGutter = (labels: readonly string[]): number => Math.max(44, widestLabel(labels, TICK_SIZE) + 10);

function drawCartesian(shapes: Shape[], hits: ChartHit[], resolved: ResolvedChart, plot: Rect): void {
  // A series set to draw as a type that can't share category axes (pie, scatter…) draws as the
  // chart's own type instead of vanishing.
  const chart: ResolvedChart = {
    ...resolved,
    series: resolved.series.map((s) =>
      isBarType(s.type) || s.type === 'line' || s.type === 'area' ? s : { ...s, type: resolved.type },
    ),
  };
  const categories = categoryCount(chart);
  if (categories === 0) return;

  const horizontal = chart.type === 'bar';
  const stacked = chart.stacking === 'stacked' || chart.stacking === 'percent';
  const percent = chart.stacking === 'percent';

  const extent = valueExtent(chart, categories, stacked, percent);
  const ticks = niceTicks(extent.min, extent.max);
  const { min, max } = ticks;
  const tickLabels = ticks.values.map((v) => formatTick(v, percent));

  const labels = Array.from({ length: categories }, (_, c) => categoryLabel(chart, c));
  let area: Rect;
  if (horizontal) {
    const widest = widestLabel(labels, CATEGORY_SIZE);
    const left = Math.min(Math.max(44, widest + 12), plot.width * 0.4);
    area = axisArea(chart, plot, left, 26);
  } else {
    area = axisArea(chart, plot, tickGutter(tickLabels), 26);
  }
  if (area.width < 10 || area.height < 10) return;

  drawValueAxis(shapes, area, ticks.values, tickLabels, min, max, horizontal);

  const bandSize = (horizontal ? area.height : area.width) / categories;
  const indexed = chart.series.map((s, i) => ({ s, i }));
  const bars = indexed.filter((e) => isBarType(e.s.type));
  const lines = indexed.filter((e) => e.s.type === 'line');
  const areas = indexed.filter((e) => e.s.type === 'area');

  // Areas first (filled, behind), then bars, then lines on top.
  const lanes = stackAreas(
    areas.map((e) => e.s),
    categories,
    stacked,
    percent,
  );
  areas.forEach((e, k) => {
    const lane = lanes[k]!;
    const topPts = lane.top.map((v, c) => categoryPoint(c, v, area, min, max, bandSize, horizontal));
    const basePts = lane.base.map((v, c) => categoryPoint(c, v, area, min, max, bandSize, horizontal));
    if (topPts.length === 0) return;
    const outline = [...topPts, ...basePts.reverse()];
    shapes.push({
      kind: 'path',
      d: polyline(outline) + 'Z',
      fill: { series: e.i },
      fillOpacity: 0.35,
      stroke: null,
    });
    shapes.push({ kind: 'path', d: polyline(topPts), fill: null, stroke: { series: e.i }, strokeWidth: 2 });
  });

  const barList = drawBars(
    shapes,
    bars.map((e) => e.s),
    bars.map((e) => e.i),
    area,
    categories,
    min,
    max,
    stacked,
    percent,
    horizontal,
    bandSize,
  );

  const lineValues = stackLines(
    lines.map((e) => e.s),
    categories,
    stacked,
    percent,
  );
  lines.forEach(({ i }, k) =>
    drawLineSeries(shapes, lineValues[k]!, i, area, categories, min, max, bandSize, horizontal),
  );

  // Hover: one target per category band, across the whole plot, so thin bars and lines are easy.
  const point = (c: number, v: number) => categoryPoint(c, v, area, min, max, bandSize, horizontal);
  const markers: { c: number; series: number; at: [number, number] }[] = [];
  areas.forEach((e, k) => {
    lanes[k]!.top.forEach((v, c) => {
      if (!Number.isNaN(valueAt(e.s, c))) markers.push({ c, series: e.i, at: point(c, v) });
    });
  });
  lines.forEach((e, k) => {
    lineValues[k]!.forEach((v, c) => {
      if (!Number.isNaN(v)) markers.push({ c, series: e.i, at: point(c, v) });
    });
  });
  for (let c = 0; c < categories; c++) {
    const start = (horizontal ? area.y : area.x) + bandSize * c;
    const centre = start + bandSize / 2;
    const band: Rect = horizontal
      ? { x: area.x, y: start, width: area.width, height: bandSize }
      : { x: start, y: area.y, width: bandSize, height: area.height };
    const highlight: Shape[] = [];
    if (lines.length + areas.length > 0) {
      highlight.push(
        horizontal
          ? {
              kind: 'line',
              x1: area.x,
              y1: r2(centre),
              x2: rectRight(area),
              y2: r2(centre),
              stroke: 'muted',
              width: 1,
            }
          : {
              kind: 'line',
              x1: r2(centre),
              y1: area.y,
              x2: r2(centre),
              y2: rectBottom(area),
              stroke: 'muted',
              width: 1,
            },
      );
    }
    // The anchor is the mark that reaches furthest up (or right, for a bar chart).
    let reach = horizontal ? -Infinity : Infinity;
    const further = (along: number) => (reach = horizontal ? Math.max(reach, along) : Math.min(reach, along));
    for (const m of markers) {
      if (m.c !== c) continue;
      highlight.push(markerHighlight(m.at[0], m.at[1], m.series));
      further(horizontal ? m.at[0] : m.at[1]);
    }
    const hoverBars: HoverBar[] = barList
      .filter((b) => b.category === c)
      .map((b) => {
        const anchor = horizontal
          ? { x: r2(b.x + b.width), y: r2(b.y + b.height / 2) }
          : { x: r2(b.x + b.width / 2), y: r2(b.y) };
        further(horizontal ? anchor.x : anchor.y);
        return {
          rect: { x: r2(b.x), y: r2(b.y), width: r2(b.width), height: r2(b.height) },
          series: b.seriesIndex,
          anchor,
        };
      });
    if (!Number.isFinite(reach)) reach = horizontal ? area.x + area.width / 2 : area.y + area.height / 2;
    hits.push({
      region: { kind: 'rect', ...band },
      info: categoryInfo(chart, c),
      anchor: horizontal ? { x: r2(reach), y: r2(centre) } : { x: r2(centre), y: r2(reach) },
      highlight,
      bars: hoverBars,
      horizontal,
    });
  }

  // Category labels, thinned when they would overlap.
  if (horizontal) {
    const step = Math.max(1, Math.ceil((CATEGORY_SIZE + 2) / bandSize));
    for (let c = 0; c < categories; c += step) {
      const label = labels[c]!;
      if (label.length === 0) continue;
      shapes.push({
        kind: 'text',
        x: area.x - 6,
        y: area.y + bandSize * (c + 0.5),
        text: label,
        size: CATEGORY_SIZE,
        paint: 'muted',
        anchor: 'end',
      });
    }
  } else {
    const widest = widestLabel(labels, CATEGORY_SIZE);
    const step = Math.max(1, Math.ceil((widest + 6) / bandSize));
    for (let c = 0; c < categories; c += step) {
      const label = labels[c]!;
      if (label.length === 0) continue;
      shapes.push({
        kind: 'text',
        x: area.x + bandSize * (c + 0.5),
        y: rectBottom(area) + 6 + CATEGORY_SIZE * 0.65,
        text: label,
        size: CATEGORY_SIZE,
        paint: 'muted',
        anchor: 'middle',
      });
    }
  }

  drawAxisTitles(shapes, chart, plot, area);
}

/** Bar rectangles in pixels, for {@link drawCartesian} and tests. */
export interface BarRect {
  readonly seriesIndex: number;
  readonly category: number;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export function barRects(
  series: readonly ResolvedSeries[],
  indices: readonly number[],
  area: Rect,
  categories: number,
  min: number,
  max: number,
  stacked: boolean,
  percent: boolean,
  horizontal: boolean,
  bandSize: number,
): BarRect[] {
  if (series.length === 0) return [];
  const out: BarRect[] = [];
  const zero = valuePos(0, min, max, area, horizontal);
  const inner = bandSize * 0.7;
  const pad = (bandSize - inner) / 2;
  const bandStart = (c: number) => (horizontal ? area.y : area.x) + bandSize * c;
  const rect = (
    si: number,
    category: number,
    offset: number,
    thickness: number,
    base: number,
    value: number,
  ): BarRect =>
    horizontal
      ? {
          seriesIndex: si,
          category,
          x: Math.min(base, value),
          y: offset,
          width: Math.abs(value - base),
          height: thickness,
        }
      : {
          seriesIndex: si,
          category,
          x: offset,
          y: Math.min(base, value),
          width: thickness,
          height: Math.abs(value - base),
        };

  if (stacked) {
    for (const seg of stackBars(series, indices, categories, percent)) {
      const p0 = valuePos(seg.start, min, max, area, horizontal);
      const p1 = valuePos(seg.end, min, max, area, horizontal);
      out.push(rect(seg.seriesIndex, seg.category, bandStart(seg.category) + pad, inner, p0, p1));
    }
    return out;
  }
  const slot = inner / series.length;
  for (let c = 0; c < categories; c++) {
    series.forEach((s, k) => {
      const v = valueAt(s, c);
      if (Number.isNaN(v)) return;
      const p1 = valuePos(v, min, max, area, horizontal);
      out.push(rect(indices[k] ?? k, c, bandStart(c) + pad + slot * k, slot * 0.9, zero, p1));
    });
  }
  return out;
}

function drawBars(
  shapes: Shape[],
  series: readonly ResolvedSeries[],
  indices: readonly number[],
  area: Rect,
  categories: number,
  min: number,
  max: number,
  stacked: boolean,
  percent: boolean,
  horizontal: boolean,
  bandSize: number,
): BarRect[] {
  const rects = barRects(series, indices, area, categories, min, max, stacked, percent, horizontal, bandSize);
  for (const b of rects) {
    shapes.push({
      kind: 'rect',
      x: r2(b.x),
      y: r2(b.y),
      width: r2(b.width),
      height: r2(b.height),
      fill: { series: b.seriesIndex },
    });
  }
  return rects;
}

/** A hovered data point: a larger marker ringed in the background colour. */
function markerHighlight(x: number, y: number, series: number, r = 5): Shape {
  return { kind: 'circle', cx: r2(x), cy: r2(y), r, fill: { series }, stroke: 'bg', strokeWidth: 2 };
}

function drawLineSeries(
  shapes: Shape[],
  values: readonly number[],
  seriesIndex: number,
  area: Rect,
  categories: number,
  min: number,
  max: number,
  bandSize: number,
  horizontal: boolean,
): void {
  // Gaps break the line into runs.
  const runs: [number, number][][] = [];
  let run: [number, number][] = [];
  const dots: [number, number][] = [];
  for (let c = 0; c < categories; c++) {
    const v = values[c] ?? Number.NaN;
    if (Number.isNaN(v)) {
      if (run.length > 0) runs.push(run);
      run = [];
      continue;
    }
    const p = categoryPoint(c, v, area, min, max, bandSize, horizontal);
    run.push(p);
    dots.push(p);
  }
  if (run.length > 0) runs.push(run);
  for (const r of runs) {
    if (r.length > 1)
      shapes.push({
        kind: 'path',
        d: polyline(r),
        fill: null,
        stroke: { series: seriesIndex },
        strokeWidth: 2,
      });
  }
  for (const [x, y] of dots)
    shapes.push({ kind: 'circle', cx: r2(x), cy: r2(y), r: 3, fill: { series: seriesIndex } });
}

function categoryPoint(
  category: number,
  value: number,
  area: Rect,
  min: number,
  max: number,
  bandSize: number,
  horizontal: boolean,
): [number, number] {
  const along = bandSize * (category + 0.5);
  const across = valuePos(value, min, max, area, horizontal);
  return horizontal ? [across, area.y + along] : [area.x + along, across];
}

function polyline(points: readonly (readonly [number, number])[]): string {
  return points.map(([x, y], i) => `${i === 0 ? 'M' : 'L'}${r2(x)} ${r2(y)}`).join('');
}

function drawValueAxis(
  shapes: Shape[],
  area: Rect,
  values: readonly number[],
  labels: readonly string[],
  min: number,
  max: number,
  horizontal: boolean,
): void {
  shapes.push({ kind: 'rect', ...area, fill: null, stroke: 'axis' });
  values.forEach((value, t) => {
    const pos = r2(valuePos(value, min, max, area, horizontal));
    const text = labels[t] ?? '';
    if (horizontal) {
      shapes.push({
        kind: 'line',
        x1: pos,
        y1: area.y,
        x2: pos,
        y2: rectBottom(area),
        stroke: 'grid',
        width: 1,
      });
      shapes.push({
        kind: 'text',
        x: pos,
        y: rectBottom(area) + 4 + TICK_SIZE * 0.65,
        text,
        size: TICK_SIZE,
        paint: 'muted',
        anchor: 'middle',
      });
    } else {
      shapes.push({
        kind: 'line',
        x1: area.x,
        y1: pos,
        x2: rectRight(area),
        y2: pos,
        stroke: 'grid',
        width: 1,
      });
      shapes.push({
        kind: 'text',
        x: area.x - 6,
        y: pos,
        text,
        size: TICK_SIZE,
        paint: 'muted',
        anchor: 'end',
      });
    }
  });
}

function drawAxisTitles(shapes: Shape[], chart: ResolvedChart, plot: Rect, area: Rect): void {
  const x = chart.xAxisTitle?.trim();
  if (x) {
    shapes.push({
      kind: 'text',
      x: area.x + area.width / 2,
      y: rectBottom(plot) - AXIS_TITLE_SIZE * 0.7,
      text: x,
      size: AXIS_TITLE_SIZE,
      paint: 'text',
      anchor: 'middle',
    });
  }
  const y = chart.yAxisTitle?.trim();
  if (y) {
    shapes.push({
      kind: 'text',
      x: plot.x + AXIS_TITLE_SIZE * 0.8,
      y: area.y + area.height / 2,
      text: y,
      size: AXIS_TITLE_SIZE,
      paint: 'text',
      anchor: 'middle',
      rotate: -90,
    });
  }
}

// ---- Pie / doughnut ---------------------------------------------------------------------------

function drawPie(
  shapes: Shape[],
  hits: ChartHit[],
  chart: ResolvedChart,
  plot: Rect,
  doughnut: boolean,
): void {
  if (!chart.series[0]) return;
  const slices = pieSlices(chart);
  if (slices.length === 0) {
    centered(shapes, 'Pie needs positive values.', plot, 12);
    return;
  }
  const radius = Math.min(plot.width, plot.height) / 2 - 12;
  if (radius <= 0) return;
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  const inner = doughnut ? radius * 0.55 : 0;
  const total = slices.reduce((sum, s) => sum + s.value, 0);
  for (const s of slices) {
    const d = wedgePath(cx, cy, radius, inner, s.start, s.end);
    shapes.push({ kind: 'path', d, fill: { series: s.index }, stroke: 'bg', strokeWidth: 1 });
    const mid = (s.start + s.end) / 2;
    const at = inner + (radius - inner) * 0.6;
    hits.push({
      region: { kind: 'wedge', cx, cy, inner, outer: radius, start: s.start, end: s.end },
      info: sliceInfo(chart, s.index, s.value, total),
      anchor: { x: r2(cx + at * Math.cos(mid)), y: r2(cy + at * Math.sin(mid)) },
      highlight: [{ kind: 'path', d, fill: null, stroke: 'text', strokeWidth: 2 }],
    });
  }
}

// ---- Radar ------------------------------------------------------------------------------------

function drawRadar(
  shapes: Shape[],
  hits: ChartHit[],
  chart: ResolvedChart,
  plot: Rect,
  filled: boolean,
): void {
  const axes = categoryCount(chart);
  if (axes < 3) {
    centered(shapes, 'Radar needs at least 3 categories.', plot, 12);
    return;
  }
  let max = -Infinity;
  for (const s of chart.series) for (const v of s.values) if (!Number.isNaN(v)) max = Math.max(max, v);
  if (!Number.isFinite(max) || max <= 0) max = 1;

  const radius = Math.min(plot.width, plot.height) / 2 - 28;
  if (radius <= 0) return;
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  const angle = (a: number) => -Math.PI / 2 + (a * Math.PI * 2) / axes;

  for (let a = 0; a < axes; a++) {
    const ang = angle(a);
    const ox = cx + radius * Math.cos(ang);
    const oy = cy + radius * Math.sin(ang);
    shapes.push({ kind: 'line', x1: r2(cx), y1: r2(cy), x2: r2(ox), y2: r2(oy), stroke: 'grid', width: 1 });
    const label = categoryLabel(chart, a);
    if (label.length > 0) {
      // Just beyond the spoke tip, anchored away from the centre so it doesn't sit on the spoke.
      const lx = cx + (radius + 12) * Math.cos(ang);
      const ly = cy + (radius + 12) * Math.sin(ang);
      const cos = Math.cos(ang);
      shapes.push({
        kind: 'text',
        x: r2(lx),
        y: r2(ly),
        text: label,
        size: TICK_SIZE,
        paint: 'muted',
        anchor: Math.abs(cos) < 0.2 ? 'middle' : cos > 0 ? 'start' : 'end',
      });
    }
  }

  const pointsBySeries: [number, number][][] = [];
  chart.series.forEach((s, si) => {
    const points: [number, number][] = [];
    pointsBySeries.push(points);
    for (let a = 0; a < axes; a++) {
      const v = valueAt(s, a);
      const mag = Number.isNaN(v) ? 0 : Math.max(0, v) / max;
      const ang = angle(a);
      points.push([cx + radius * mag * Math.cos(ang), cy + radius * mag * Math.sin(ang)]);
    }
    shapes.push({
      kind: 'path',
      d: polyline(points) + 'Z',
      fill: filled ? { series: si } : null,
      fillOpacity: filled ? 0.3 : undefined,
      stroke: { series: si },
      strokeWidth: 2,
    });
  });

  // Hover: the spoke nearest the pointer, out to just past its label.
  const half = Math.PI / axes;
  const info = { ...chart, stacking: 'none' as const };
  for (let a = 0; a < axes; a++) {
    const ang = angle(a);
    const tip: [number, number] = [cx + radius * Math.cos(ang), cy + radius * Math.sin(ang)];
    const highlight: Shape[] = [
      { kind: 'line', x1: r2(cx), y1: r2(cy), x2: r2(tip[0]), y2: r2(tip[1]), stroke: 'muted', width: 1 },
    ];
    let anchor = tip;
    let reach = -1;
    chart.series.forEach((s, si) => {
      if (Number.isNaN(valueAt(s, a))) return;
      const p = pointsBySeries[si]![a]!;
      highlight.push(markerHighlight(p[0], p[1], si, 4));
      const d = Math.hypot(p[0] - cx, p[1] - cy);
      if (d > reach) {
        reach = d;
        anchor = p;
      }
    });
    hits.push({
      region: { kind: 'wedge', cx, cy, inner: 0, outer: radius + 20, start: ang - half, end: ang + half },
      info: categoryInfo(info, a),
      anchor: { x: r2(anchor[0]), y: r2(anchor[1]) },
      highlight,
    });
  }
}

// ---- Scatter / bubble -------------------------------------------------------------------------

/** The (x, y) of point `i`: its X value, or its 1-based index when the series has no X values. */
function xyAt(s: ResolvedSeries, i: number): [number, number] {
  const y = s.values[i] ?? NaN;
  const x = s.xValues ? (i < s.xValues.length ? s.xValues[i]! : i + 1) : i + 1;
  return [x, y];
}

function drawScatter(
  shapes: Shape[],
  hits: ChartHit[],
  chart: ResolvedChart,
  plot: Rect,
  bubble: boolean,
): void {
  let xMin = Infinity;
  let xMax = -Infinity;
  let yMin = Infinity;
  let yMax = -Infinity;
  for (const s of chart.series) {
    for (let i = 0; i < s.values.length; i++) {
      const [x, y] = xyAt(s, i);
      if (Number.isNaN(x) || Number.isNaN(y)) continue;
      xMin = Math.min(xMin, x);
      xMax = Math.max(xMax, x);
      yMin = Math.min(yMin, y);
      yMax = Math.max(yMax, y);
    }
  }
  if (xMin > xMax) return;
  yMin = Math.min(yMin, 0);

  const xt = niceTicks(xMin, xMax);
  const yt = niceTicks(yMin, yMax);
  const yLabels = yt.values.map((v) => formatNumber(v));
  const area = axisArea(chart, plot, tickGutter(yLabels), 26);
  if (area.width < 10 || area.height < 10) return;

  drawValueAxis(shapes, area, yt.values, yLabels, yt.min, yt.max, false);
  for (const v of xt.values) {
    const px = r2(valuePos(v, xt.min, xt.max, area, true));
    shapes.push({
      kind: 'line',
      x1: px,
      y1: rectBottom(area),
      x2: px,
      y2: rectBottom(area) + 3,
      stroke: 'axis',
      width: 1,
    });
    shapes.push({
      kind: 'text',
      x: px,
      y: rectBottom(area) + 4 + TICK_SIZE * 0.65,
      text: formatNumber(v),
      size: TICK_SIZE,
      paint: 'muted',
      anchor: 'middle',
    });
  }

  let maxSize = 1;
  if (bubble) {
    let m = -Infinity;
    for (const s of chart.series)
      for (const v of s.bubbleSizes ?? []) if (!Number.isNaN(v) && v > 0) m = Math.max(m, v);
    if (Number.isFinite(m)) maxSize = m;
  }

  chart.series.forEach((s, si) => {
    for (let i = 0; i < s.values.length; i++) {
      const [x, y] = xyAt(s, i);
      if (Number.isNaN(x) || Number.isNaN(y)) continue;
      const px = valuePos(x, xt.min, xt.max, area, true);
      const py = valuePos(y, yt.min, yt.max, area, false);
      let r = 4;
      const size = s.bubbleSizes?.[i];
      if (bubble && size !== undefined && !Number.isNaN(size) && size > 0)
        r = 4 + 18 * Math.sqrt(size / maxSize);
      shapes.push({
        kind: 'circle',
        cx: r2(px),
        cy: r2(py),
        r: r2(r),
        fill: { series: si },
        opacity: bubble ? 0.55 : 1,
      });
      hits.push({
        region: { kind: 'point', x: r2(px), y: r2(py), radius: Math.max(POINT_HIT_RADIUS, r) },
        info: pointInfo(chart, si, x, y, bubble ? (size ?? NaN) : null),
        anchor: { x: r2(px), y: r2(py - r) },
        highlight: [
          {
            kind: 'circle',
            cx: r2(px),
            cy: r2(py),
            r: r2(r + 2),
            fill: null,
            stroke: 'text',
            strokeWidth: 2,
          },
        ],
      });
    }
  });

  drawAxisTitles(shapes, chart, plot, area);
}

/** The widest label's estimated width; a loop, since spreading 100k+ labels overflows the stack. */
function widestLabel(labels: readonly string[], size: number): number {
  let widest = 0;
  for (const l of labels) widest = Math.max(widest, estimateTextWidth(l, size));
  return widest;
}
