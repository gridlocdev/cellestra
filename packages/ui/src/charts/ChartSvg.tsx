import type { ResolvedChart } from '@cellestra/core';
import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'preact/hooks';
import { type Hover, hitTest, type Point, placeTooltip } from './chartHover.ts';
import { seriesColor } from './chartMath.ts';
import { nextFrame } from '../schedule.ts';
import { buildChartScene, type ChartScene, type Paint, type Shape } from './chartScene.ts';

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
  /** Show hover tooltips and highlights (the chart sheet); off for static copies. */
  readonly interactive?: boolean;
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
          {...paintProps(shape.fill, shape.stroke)}
          fill-opacity={shape.opacity}
          stroke-width={shape.stroke ? (shape.strokeWidth ?? 1) : undefined}
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

/** The tooltip for `hover`, placed by {@link placeTooltip} once its size is known. */
function ChartTooltip(props: { hover: Hover; box: { width: number; height: number }; themeClass: string }) {
  const ref = useRef<HTMLDivElement>(null);
  const [at, setAt] = useState<Point | null>(null);
  const { hover, box } = props;
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    setAt(placeTooltip(hover.anchor, { width: el.offsetWidth, height: el.offsetHeight }, box));
  }, [hover, box.width, box.height]);
  const { info } = hover;
  return (
    <div
      ref={ref}
      class={`cs-chart-tip${props.themeClass}`}
      style={at ? { left: `${at.x}px`, top: `${at.y}px` } : { left: '0px', top: '0px', visibility: 'hidden' }}
    >
      {info.title !== '' && <div class="cs-chart-tip-title">{info.title}</div>}
      {info.rows.map((row, i) => (
        <div key={i} class={row.emphasis ? 'cs-chart-tip-row cs-chart-tip-on' : 'cs-chart-tip-row'}>
          <span
            class="cs-chart-tip-swatch"
            style={
              row.series === undefined ? { visibility: 'hidden' } : { background: seriesColor(row.series) }
            }
          />
          <span class="cs-chart-tip-label">{row.label}</span>
          <span class="cs-chart-tip-value">{row.value}</span>
        </div>
      ))}
      {info.note && <div class="cs-chart-tip-note">{info.note}</div>}
    </div>
  );
}

/**
 * Pointer tracking for an interactive chart: the latest pointer position is hit-tested at most
 * once per animation frame.
 */
function useChartHover(scene: ChartScene) {
  const [hover, setHover] = useState<Hover | null>(null);
  const cancelFrame = useRef<(() => void) | null>(null);
  const pending = useRef<Point | null>(null);
  const sceneRef = useRef(scene);
  sceneRef.current = scene;
  useEffect(() => () => cancelFrame.current?.(), []);
  useEffect(() => setHover(null), [scene]);

  const track = (e: PointerEvent) => {
    const rect = (e.currentTarget as Element).getBoundingClientRect();
    pending.current = { x: e.clientX - rect.left, y: e.clientY - rect.top };
    if (cancelFrame.current) return;
    cancelFrame.current = nextFrame(() => {
      cancelFrame.current = null;
      const p = pending.current;
      if (p) setHover(hitTest(sceneRef.current, p.x, p.y));
    });
  };
  const leave = (e: PointerEvent) => {
    // A touch lifts with a leave; keep its tooltip until the next tap.
    if (e.pointerType === 'touch') return;
    pending.current = null;
    setHover(null);
  };
  return { hover, onPointerMove: track, onPointerDown: track, onPointerLeave: leave };
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
  if (!props.interactive) return <ChartSvgImage scene={scene} chart={props.chart} theme={theme} />;
  return <InteractiveChart scene={scene} chart={props.chart} theme={theme} />;
}

function InteractiveChart(props: { scene: ChartScene; chart: ResolvedChart | null; theme: string }) {
  const { hover, ...handlers } = useChartHover(props.scene);
  const { width, height } = props.scene;
  const box = useMemo(() => ({ width, height }), [width, height]);
  return (
    <div class="cs-chart-frame" style={{ width: `${width}px`, height: `${height}px` }}>
      <ChartSvgImage
        scene={props.scene}
        chart={props.chart}
        theme={props.theme}
        highlight={hover?.highlight}
        handlers={handlers}
      />
      {hover && <ChartTooltip hover={hover} box={box} themeClass={props.theme} />}
    </div>
  );
}

interface PointerHandlers {
  readonly onPointerMove: (e: PointerEvent) => void;
  readonly onPointerDown: (e: PointerEvent) => void;
  readonly onPointerLeave: (e: PointerEvent) => void;
}

function ChartSvgImage(props: {
  scene: ChartScene;
  chart: ResolvedChart | null;
  theme: string;
  highlight?: readonly Shape[];
  handlers?: PointerHandlers;
}) {
  const { scene, theme } = props;
  const { width, height } = scene;
  return (
    <svg
      class={`cs-chart${theme}`}
      width={width}
      height={height}
      viewBox={`0 0 ${width} ${height}`}
      role="img"
      aria-label={props.chart?.title || 'Chart'}
      xmlns="http://www.w3.org/2000/svg"
      {...props.handlers}
    >
      {scene.shapes.map(renderShape)}
      {props.highlight && <g class="cs-chart-highlight">{props.highlight.map(renderShape)}</g>}
    </svg>
  );
}
