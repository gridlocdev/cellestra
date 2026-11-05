import type { ChartType } from '@cellestra/charts';
import type { ResolvedChart, ResolvedSeries } from '@cellestra/core';
import { describe, expect, it } from 'vitest';
import {
  formatNumber,
  formatTick,
  layoutLegend,
  legendItems,
  niceStep,
  niceTicks,
  pieSlices,
  seriesColor,
  SERIES_PALETTE,
  stackAreas,
  stackBars,
  valueExtent,
  valuePos,
  wedgePath,
} from '../src/charts/chartMath.ts';
import { barRects, buildChartScene, type Shape } from '../src/charts/chartScene.ts';
import { chartFamily } from '../src/charts/ChartSheetView.tsx';

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
    showLegend: true,
    xAxisTitle: null,
    yAxisTitle: null,
    categories: [],
    series,
    isEmpty: series.every((s) => s.values.length === 0),
    ...init,
  };
}

describe('niceTicks', () => {
  it('picks 1/2/2.5/5 × 10ⁿ steps', () => {
    expect(niceStep(0.7)).toBe(1);
    expect(niceStep(1.3)).toBe(2);
    expect(niceStep(2.2)).toBe(2.5);
    expect(niceStep(3)).toBe(5);
    expect(niceStep(7)).toBe(10);
    expect(niceStep(0.03)).toBe(0.05);
    expect(niceStep(1700)).toBe(2000);
  });

  it('covers the extent on round bounds', () => {
    expect(niceTicks(0, 87)).toEqual({ min: 0, max: 100, step: 20, values: [0, 20, 40, 60, 80, 100] });
    const t = niceTicks(-13, 42);
    expect(t.step).toBe(20);
    expect(t.min).toBe(-20);
    expect(t.max).toBe(60);
  });

  it('a 0..1 percent axis steps by 0.2 without float noise', () => {
    expect(niceTicks(0, 1).values).toEqual([0, 0.2, 0.4, 0.6, 0.8, 1]);
  });

  it('a flat or reversed extent still yields ticks', () => {
    expect(niceTicks(5, 5).values.length).toBeGreaterThan(1);
    expect(niceTicks(10, 0).min).toBe(0);
  });
});

describe('formatting', () => {
  it('ports "#,##0" and "0.##"', () => {
    expect(formatNumber(1234567.4)).toBe('1,234,567');
    expect(formatNumber(-2500)).toBe('-2,500');
    expect(formatNumber(3.14159)).toBe('3.14');
    expect(formatNumber(0.1 + 0.2)).toBe('0.3');
    expect(formatNumber(-0.001)).toBe('0');
    expect(formatTick(0.6, true)).toBe('60%');
  });

  it('cycles the palette', () => {
    expect(seriesColor(SERIES_PALETTE.length)).toBe(seriesColor(0));
  });
});

describe('valueExtent', () => {
  const c = chart([S([5, -3, 2]), S([4, 1, NaN])]);

  it('grouped bars span the raw values and include zero', () => {
    expect(valueExtent(c, 3, false, false)).toEqual({ min: -3, max: 5 });
  });

  it('stacked bars span their positive and negative sums', () => {
    expect(valueExtent(c, 3, true, false)).toEqual({ min: -3, max: 9 });
  });

  it('percent spans 0..1', () => {
    expect(valueExtent(c, 3, true, true)).toEqual({ min: 0, max: 1 });
  });

  it('stacked areas span their running positive sum; lines stay raw', () => {
    const mixed = chart([S([1, 2], 'area'), S([3, 4], 'area'), S([10, 0], 'line')]);
    expect(valueExtent(mixed, 2, true, false)).toEqual({ min: 0, max: 10 });
    expect(valueExtent(mixed, 2, false, false)).toEqual({ min: 0, max: 10 });
    expect(valueExtent(chart([S([1, 2], 'area'), S([3, 5], 'area')]), 2, true, false).max).toBe(7);
  });

  it('an all-zero chart still has a non-empty extent', () => {
    expect(valueExtent(chart([S([0, 0])]), 2, false, false)).toEqual({ min: 0, max: 1 });
  });
});

describe('stacking', () => {
  const a = S([2, 1, NaN]);
  const b = S([3, -1, 4]);

  it('stacks in series order, positives up and negatives down, skipping gaps', () => {
    expect(stackBars([a, b], [0, 1], 3, false)).toEqual([
      { seriesIndex: 0, category: 0, start: 0, end: 2 },
      { seriesIndex: 1, category: 0, start: 2, end: 5 },
      { seriesIndex: 0, category: 1, start: 0, end: 1 },
      { seriesIndex: 1, category: 1, start: 0, end: -1 },
      { seriesIndex: 1, category: 2, start: 0, end: 4 },
    ]);
  });

  it('100% stacking scales each category to 1', () => {
    const segs = stackBars([a, b], [0, 1], 1, true);
    expect(segs.map((s) => [s.start, s.end])).toEqual([
      [0, 0.4],
      [0.4, 1],
    ]);
  });

  it('stacked areas sit on the lanes beneath them', () => {
    const lanes = stackAreas([S([1, 2]), S([3, NaN])], 2, true, false);
    expect(lanes[0]).toEqual({ base: [0, 0], top: [1, 2] });
    expect(lanes[1]).toEqual({ base: [1, 2], top: [4, 2] });
    const pct = stackAreas([S([1, 1]), S([3, 1])], 2, true, true);
    expect(pct[1]).toEqual({ base: [0.25, 0.5], top: [1, 1] });
    const flat = stackAreas([S([1, -2])], 2, false, false);
    expect(flat[0]).toEqual({ base: [0, 0], top: [1, -2] });
  });
});

describe('layout', () => {
  const area = { x: 10, y: 20, width: 100, height: 200 };

  it('maps values onto the axis', () => {
    expect(valuePos(0, 0, 10, area, false)).toBe(220);
    expect(valuePos(10, 0, 10, area, false)).toBe(20);
    expect(valuePos(5, 0, 10, area, true)).toBe(60);
  });

  it('grouped bars share 70% of each band, 90% of each slot', () => {
    const rects = barRects([S([10]), S([5])], [0, 1], area, 1, 0, 10, false, false, false, 100);
    expect(rects).toHaveLength(2);
    const [r0, r1] = rects;
    expect(r0).toMatchObject({ seriesIndex: 0, x: 25, y: 20, height: 200 });
    expect(r0!.width).toBeCloseTo(31.5);
    expect(r1!.x).toBeCloseTo(60);
    expect(r1).toMatchObject({ y: 120, height: 100 });
  });

  it('horizontal bars grow from the zero line rightwards', () => {
    const [r] = barRects([S([5])], [0], area, 1, 0, 10, false, false, true, 200);
    expect(r).toMatchObject({ x: 10, width: 50, y: 50, height: 126 }); // 70% band, 90% of the slot
  });

  it('pie slices start at 12 o’clock and skip non-positive values', () => {
    const slices = pieSlices(chart([S([1, 0, 3, NaN])], { type: 'pie' }));
    expect(slices.map((s) => s.index)).toEqual([0, 2]);
    expect(slices[0]!.start).toBeCloseTo(-Math.PI / 2);
    expect(slices[1]!.end).toBeCloseTo(Math.PI * 1.5);
  });

  it('a lone full-circle wedge is split into two arcs', () => {
    const d = wedgePath(50, 50, 40, 0, -Math.PI / 2, Math.PI * 1.5);
    expect(d.match(/A/g)).toHaveLength(2);
  });

  it('pie legends name slices; others name series', () => {
    const pie = chart([S([1, 0, 2])], { type: 'pie', categories: ['a', 'b', 'c'] });
    expect(legendItems(pie)).toEqual([
      { name: 'a', index: 0 },
      { name: 'c', index: 2 },
    ]);
    expect(legendItems(chart([S([1], 'line', 'x'), S([2], 'line', 'y')])).map((i) => i.name)).toEqual([
      'x',
      'y',
    ]);
  });

  it('legend rows centre and wrap', () => {
    const one = layoutLegend([{ name: 'abc', index: 0 }], 400);
    expect(one.height).toBe(26);
    expect(one.items[0]!.x).toBeGreaterThan(150);
    const many = layoutLegend(
      Array.from({ length: 12 }, (_, i) => ({ name: `Series number ${i}`, index: i })),
      300,
    );
    expect(many.height).toBeGreaterThan(26);
    expect(Math.max(...many.items.map((i) => i.row))).toBe(many.height / 26 - 1);
  });
});

describe('buildChartScene', () => {
  const texts = (shapes: readonly Shape[]) => shapes.flatMap((s) => (s.kind === 'text' ? [s.text] : []));
  const types: ChartType[] = [
    'column',
    'bar',
    'line',
    'area',
    'pie',
    'doughnut',
    'scatter',
    'bubble',
    'radar',
    'filledRadar',
  ];

  it('draws every chart type without throwing', () => {
    for (const type of types) {
      const c = chart([S([1, 4, 2, 5], type, 'A'), S([3, 1, 2, 0], type, 'B')], {
        type,
        title: 'T',
        categories: ['w', 'x', 'y', 'z'],
      });
      const scene = buildChartScene(c, 600, 400);
      expect(scene.shapes.length, type).toBeGreaterThan(5);
      expect(texts(scene.shapes), type).toContain('T');
    }
  });

  it('shows placeholders for no chart and no data', () => {
    expect(texts(buildChartScene(null, 300, 200).shapes)).toEqual(['No chart']);
    expect(texts(buildChartScene(chart([]), 300, 200).shapes)[0]).toMatch(/^No data/);
  });

  it('labels a percent axis and draws axis titles', () => {
    const c = chart([S([1, 2]), S([1, 2])], {
      stacking: 'percent',
      xAxisTitle: 'Month',
      yAxisTitle: 'Share',
    });
    const t = texts(buildChartScene(c, 600, 400).shapes);
    expect(t).toEqual(expect.arrayContaining(['0%', '100%', 'Month', 'Share']));
  });

  it('radar with fewer than three spokes explains itself', () => {
    const c = chart([S([1, 2], 'radar')], { type: 'radar' });
    expect(texts(buildChartScene(c, 600, 400).shapes)).toContain('Radar needs at least 3 categories.');
  });

  it('maps chart types to panel families', () => {
    expect(chartFamily('doughnut')).toBe('partToWhole');
    expect(chartFamily('bubble')).toBe('xy');
    expect(chartFamily('filledRadar')).toBe('radar');
    expect(chartFamily('area')).toBe('cartesian');
  });
});

describe('niceTicks at the edges of the double range', () => {
  it('never returns NaN or loops on huge spans', () => {
    const t = niceTicks(-1e308, 1e308);
    expect(t.values.every(Number.isFinite)).toBe(true);
    expect(t.values.length).toBeLessThanOrEqual(1001);
  });

  it('handles a span below the precision of its magnitude', () => {
    const t = niceTicks(1e20, 1e20 + 1);
    expect(t.values.every(Number.isFinite)).toBe(true);
  });
});
