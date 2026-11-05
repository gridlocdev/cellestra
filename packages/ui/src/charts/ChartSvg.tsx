import type { ResolvedChart } from '@cellestra/core';
import { useMemo } from 'preact/hooks';
import { seriesColor } from './chartMath.ts';
import { buildChartScene, type Paint, type Shape } from './chartScene.ts';

export interface ChartSvgProps {
  /** The chart to draw; null shows a "No chart" placeholder. */
  readonly chart: ResolvedChart | null;
  readonly width: number;
  readonly height: number;
  /**
   * Force the dark (true) or light (false) chart theme. Omitted, the chart follows the app's theme
   * variables (`--cs-bg`, `--cs-text`, …).
   */
  readonly dark?: boolean;
}

/** Theme roles become CSS classes (see charts.css); series indices become palette colours. */
function fillProps(paint: Paint | null): { fill?: string; class?: string } {
  if (paint === null) return { fill: 'none' };
  if (typeof paint === 'object') return { fill: seriesColor(paint.series) };
  return { class: `cs-chart-fill-${paint}` };
}

function strokeProps(paint: Paint | null | undefined): { stroke?: string; class?: string } {
  if (!paint) return {};
  if (typeof paint === 'object') return { stroke: seriesColor(paint.series) };
  return { class: `cs-chart-stroke-${paint}` };
}

function paintProps(fill: Paint | null, stroke: Paint | null | undefined) {
  const f = fillProps(fill);
  const s = strokeProps(stroke);
  const cls = [f.class, s.class].filter(Boolean).join(' ');
  return { fill: f.fill, stroke: s.stroke, class: cls || undefined };
}

function renderShape(shape: Shape, key: number) {
  switch (shape.kind) {
    case 'rect':
      return (
        <rect
          key={key}
          x={shape.x}
          y={shape.y}
          width={shape.width}
          height={shape.height}
          {...paintProps(shape.fill, shape.stroke)}
          stroke-width={shape.stroke ? 1 : undefined}
          shape-rendering={shape.stroke ? 'crispEdges' : undefined}
        />
      );
    case 'line':
      return (
        <line
          key={key}
          x1={shape.x1}
          y1={shape.y1}
          x2={shape.x2}
          y2={shape.y2}
          {...strokeProps(shape.stroke)}
          stroke-width={shape.width}
          shape-rendering="crispEdges"
        />
      );
    case 'path':
      return (
        <path
          key={key}
          d={shape.d}
          {...paintProps(shape.fill, shape.stroke)}
          fill-opacity={shape.fillOpacity}
          stroke-width={shape.stroke ? (shape.strokeWidth ?? 1) : undefined}
          stroke-linejoin="round"
        />
      );
    case 'circle':
      return (
        <circle
          key={key}
          cx={shape.cx}
          cy={shape.cy}
          r={shape.r}
          {...fillProps(shape.fill)}
          fill-opacity={shape.opacity}
        />
      );
    case 'text':
      return (
        <text
          key={key}
          x={shape.x}
          y={shape.y}
          font-size={shape.size}
          font-weight={shape.bold ? 600 : undefined}
          text-anchor={shape.anchor}
          dominant-baseline="central"
          transform={shape.rotate ? `rotate(${shape.rotate} ${shape.x} ${shape.y})` : undefined}
          {...fillProps(shape.paint)}
        >
          {shape.text}
        </text>
      );
  }
}

/**
 * A from-scratch SVG chart renderer for every Cellestra chart type (column, bar, line, area, pie,
 * doughnut, scatter, bubble, radar, filled radar), with grouped/stacked/100% variants and
 * per-series type overrides for combo charts. Layout lives in {@link buildChartScene}.
 */
export function ChartSvg(props: ChartSvgProps) {
  const width = Math.max(0, Math.floor(props.width));
  const height = Math.max(0, Math.floor(props.height));
  const scene = useMemo(() => buildChartScene(props.chart, width, height), [props.chart, width, height]);
  const theme = props.dark === undefined ? '' : props.dark ? ' cs-chart-dark' : ' cs-chart-light';
  return (
    <svg
      class={`cs-chart${theme}`}
      width={width}
      height={height}
      viewBox={`0 0 ${width} ${height}`}
      role="img"
      aria-label={props.chart?.title || 'Chart'}
      xmlns="http://www.w3.org/2000/svg"
    >
      {scene.shapes.map(renderShape)}
    </svg>
  );
}
