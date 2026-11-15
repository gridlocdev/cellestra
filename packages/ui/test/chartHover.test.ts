import type { ChartType } from '@cellestra/charts';
import type { ResolvedChart, ResolvedSeries } from '@cellestra/core';
import { describe, expect, it } from 'vitest';
import { type ChartHit, type Hover, hitTest, placeTooltip } from '../src/charts/chartHover.ts';
import { buildChartScene, type ChartScene } from '../src/charts/chartScene.ts';

const S = (values: number[], type: ChartType = 'column', name = 's'): ResolvedSeries => ({
  name,
  type,
  values,
  xValues: null,
  bubbleSizes: null,
});

function chart(series: ResolvedSeries[], init: Partial<ResolvedChart> = {}): ResolvedChart {
  return {
    title: '',
    type: series[0]?.type ?? 'column',
    stacking: 'none',
    showLegend: false,
    xAxisTitle: null,
    yAxisTitle: null,
    categories: [],
    series,
    isEmpty: false,
    ...init,
  };
}

const scene = (c: ResolvedChart): ChartScene => buildChartScene(c, 600, 400);

/** The centre of hit `i`'s region (a band's middle, a point itself). */
function centreOf(hit: ChartHit): { x: number; y: number } {
  const r = hit.region;
  if (r.kind === 'rect') return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
  if (r.kind === 'point') return { x: r.x, y: r.y };
  const mid = (r.start + r.end) / 2;
  const d = (r.inner + r.outer) / 2;
  return { x: r.cx + d * Math.cos(mid), y: r.cy + d * Math.sin(mid) };
}

const values = (h: Hover | null) => h?.info.rows.map((r) => [r.label, r.value]);
const emphasised = (h: Hover | null) => h?.info.rows.filter((r) => r.emphasis).map((r) => r.label);

describe('hover on column charts', () => {
  const grouped = chart([S([10, 4], 'column', 'North'), S([2, NaN], 'column', 'South')], {
    categories: ['Q1', 'Q2'],
    aggregate: 'sum',
    rowCounts: [3, 1],
  });

  it('a category band lists every series and emphasises the bar under the pointer', () => {
    const sc = scene(grouped);
    expect(sc.hits).toHaveLength(2);
    const bar = sc.hits[0]!.bars![0]!;
    const h = hitTest(sc, bar.rect.x + bar.rect.width / 2, bar.rect.y + bar.rect.height / 2);
    expect(h?.info.title).toBe('Q1');
    expect(values(h)).toEqual([
      ['North', '10'],
      ['South', '2'],
    ]);
    expect(emphasised(h)).toEqual(['North']);
    expect(h?.anchor).toEqual(bar.anchor);
    expect(h?.highlight.some((s) => s.kind === 'path' && s.stroke === 'text')).toBe(true);
  });

  it('the whole band answers, picking the bar whose slot holds the pointer', () => {
    const sc = scene(grouped);
    const south = sc.hits[0]!.bars!.find((b) => b.series === 1)!;
    const band = sc.hits[0]!.region;
    if (band.kind !== 'rect') throw new Error('band');
    // Well above the short South bar, but inside its slot.
    const h = hitTest(sc, south.rect.x + south.rect.width / 2, band.y + 2);
    expect(emphasised(h)).toEqual(['South']);
  });

  it('notes how many rows a grouped category merged, and only then', () => {
    const sc = scene(grouped);
    expect(hitTest(sc, centreOf(sc.hits[0]!).x, centreOf(sc.hits[0]!).y)?.info.note).toBe('Sum of 3 rows');
    expect(hitTest(sc, centreOf(sc.hits[1]!).x, centreOf(sc.hits[1]!).y)?.info.note).toBeUndefined();
    const avg = scene({ ...grouped, aggregate: 'average' });
    expect(avg.hits[0]!.info.note).toBe('Average of 3 rows');
  });

  it('shows a gap as a dash', () => {
    const sc = scene(grouped);
    expect(values(hitTest(sc, centreOf(sc.hits[1]!).x, centreOf(sc.hits[1]!).y))).toEqual([
      ['North', '4'],
      ['South', '—'],
    ]);
  });

  it('stacked bars add the total, 100% stacked bars each share', () => {
    const base = chart([S([10], 'column', 'A'), S([30], 'column', 'B')], { categories: ['X'] });
    expect(scene({ ...base, stacking: 'stacked' }).hits[0]!.info.rows.map((r) => r.value)).toEqual([
      '10',
      '30',
      '40',
    ]);
    const percent = scene({ ...base, stacking: 'percent' }).hits[0]!.info.rows;
    expect(percent.map((r) => [r.label, r.value])).toEqual([
      ['A', '10 (25%)'],
      ['B', '30 (75%)'],
      ['Total', '40'],
    ]);
  });

  it('horizontal bar charts band along y', () => {
    const sc = scene(chart([S([1, 2, 3], 'bar')], { type: 'bar', categories: ['a', 'b', 'c'] }));
    const r = sc.hits[2]!.region;
    if (r.kind !== 'rect') throw new Error('band');
    expect(r.height).toBeLessThan(r.width);
    expect(hitTest(sc, r.x + 1, r.y + r.height / 2)?.info.title).toBe('c');
  });
});

describe('hover on line charts', () => {
  const c = chart([S([1, 5, 2], 'line', 'Sales'), S([3, 3, 3], 'line', 'Cost')], {
    type: 'line',
    categories: ['Jan', 'Feb', 'Mar'],
  });

  it('the nearest category answers anywhere in the plot, with a guide and both markers', () => {
    const sc = scene(c);
    const band = sc.hits[1]!.region;
    if (band.kind !== 'rect') throw new Error('band');
    const h = hitTest(sc, band.x + band.width * 0.9, band.y + 1);
    expect(h?.info.title).toBe('Feb');
    expect(values(h)).toEqual([
      ['Sales', '5'],
      ['Cost', '3'],
    ]);
    expect(h?.highlight.filter((s) => s.kind === 'line')).toHaveLength(1);
    expect(h?.highlight.filter((s) => s.kind === 'circle')).toHaveLength(2);
    // The tooltip points at the higher of the two markers.
    const ys = h!.highlight.flatMap((s) => (s.kind === 'circle' ? [s.cy] : []));
    expect(h?.anchor.y).toBe(Math.min(...ys));
  });

  it('is nothing outside the plot area', () => {
    expect(hitTest(scene(c), 1, 1)).toBeNull();
  });
});

describe('hover on pie charts', () => {
  it('a slice gives its value and share of the total', () => {
    const c = chart([S([1, 3], 'pie', 'Votes')], { type: 'pie', categories: ['Yes', 'No'] });
    const sc = scene(c);
    const h = hitTest(sc, centreOf(sc.hits[1]!).x, centreOf(sc.hits[1]!).y);
    expect(h?.info.title).toBe('No');
    expect(values(h)).toEqual([
      ['Votes', '3'],
      ['Share', '75%'],
    ]);
    expect(hitTest(sc, 2, 200)).toBeNull();
  });

  it('a doughnut hole is not a slice', () => {
    const sc = scene(chart([S([1, 3], 'doughnut')], { type: 'doughnut' }));
    const r = sc.hits[0]!.region;
    if (r.kind !== 'wedge') throw new Error('wedge');
    expect(hitTest(sc, r.cx, r.cy)).toBeNull();
  });
});

describe('hover on scatter charts', () => {
  const c = chart([{ ...S([10, 20], 'scatter', 'Trees'), xValues: [1, 2] }], {
    type: 'scatter',
    xAxisTitle: 'Age',
  });

  it('the nearest point within reach gives X (by axis title) and Y', () => {
    const sc = scene(c);
    const p = sc.hits[1]!.region;
    if (p.kind !== 'point') throw new Error('point');
    const h = hitTest(sc, p.x + 6, p.y + 6);
    expect(h?.info.title).toBe('Trees');
    expect(values(h)).toEqual([
      ['Age', '2'],
      ['Y', '20'],
    ]);
  });

  it('nothing beyond the radius', () => {
    const sc = scene(c);
    const p = sc.hits[1]!.region;
    if (p.kind !== 'point') throw new Error('point');
    expect(hitTest(sc, p.x + 20, p.y + 20)).toBeNull();
  });

  it('bubbles add their size and skip gaps', () => {
    const b = chart([{ ...S([10, NaN], 'bubble', 'B'), xValues: [1, 2], bubbleSizes: [5, 6] }], {
      type: 'bubble',
    });
    const sc = scene(b);
    expect(sc.hits).toHaveLength(1);
    expect(sc.hits[0]!.info.rows.map((r) => [r.label, r.value])).toContainEqual(['Size', '5']);
  });
});

describe('hover on radar charts', () => {
  it('the nearest spoke lists every series there', () => {
    const c = chart([S([1, 2, 3], 'radar', 'A'), S([3, 2, 1], 'radar', 'B')], {
      type: 'radar',
      categories: ['x', 'y', 'z'],
    });
    const sc = scene(c);
    const r = sc.hits[0]!.region;
    if (r.kind !== 'wedge') throw new Error('wedge');
    // Straight up from the centre is the first spoke.
    const h = hitTest(sc, r.cx + 3, r.cy - 40);
    expect(h?.info.title).toBe('x');
    expect(values(h)).toEqual([
      ['A', '1'],
      ['B', '3'],
    ]);
  });
});

describe('placeTooltip', () => {
  const box = { width: 300, height: 200 };
  const tip = { width: 100, height: 40 };

  it('centres above the anchor', () => {
    expect(placeTooltip({ x: 150, y: 100 }, tip, box)).toEqual({ x: 100, y: 50 });
  });

  it('flips below near the top and stays inside the sides', () => {
    expect(placeTooltip({ x: 10, y: 20 }, tip, box)).toEqual({ x: 4, y: 30 });
    expect(placeTooltip({ x: 295, y: 100 }, tip, box).x).toBe(196);
  });
});
