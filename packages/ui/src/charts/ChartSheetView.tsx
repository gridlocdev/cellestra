import type {
  ChartAggregate,
  ChartDefinition,
  ChartSeries,
  ChartStacking,
  ChartType,
} from '@cellestra/charts';
import { cloneChart, createChart, createSeries, ChartTypes } from '@cellestra/charts';
import {
  ChartDataResolver,
  ChartTableBinding,
  type ChartField,
  type ChartSeriesFieldKind,
  type Workbook,
  type Worksheet,
} from '@cellestra/core';
import type { ComponentChildren } from 'preact';
import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'preact/hooks';
import { ChartSvg } from './ChartSvg.tsx';
import { ChartTableFields } from './ChartTableFields.tsx';

export interface ChartSheetViewProps {
  readonly workbook: Workbook;
  /** The chart sheet; its `chart` is the definition shown and edited. */
  readonly sheet: Worksheet;
  /** Bumped whenever the workbook changes, so the chart re-resolves its ranges. */
  readonly version: number;
  /** Force the dark/light chart theme; omitted, it follows the app theme. */
  readonly dark?: boolean;
  /** A committed edit: the host applies it (e.g. via `ChartEditCommands.setChart`). */
  readonly onChange: (next: ChartDefinition) => void;
  /**
   * "Pick on sheet" for a reference field: the host shows the data sheet, lets the user select
   * cells, and sets the field. Omitted, the buttons don't show.
   */
  readonly onPick?: (field: ChartField) => void;
}

/** The data shape a chart type needs; drives which fields the panel shows. */
export type ChartFamily = 'cartesian' | 'radar' | 'partToWhole' | 'xy';

export function chartFamily(type: ChartType): ChartFamily {
  switch (type) {
    case 'pie':
    case 'doughnut':
      return 'partToWhole';
    case 'scatter':
    case 'bubble':
      return 'xy';
    case 'radar':
    case 'filledRadar':
      return 'radar';
    case 'area':
    case 'bar':
    case 'column':
    case 'line':
    default:
      return 'cartesian';
  }
}

export const CHART_TYPE_LABELS: readonly (readonly [ChartType, string])[] = [
  ['column', 'Column'],
  ['bar', 'Bar'],
  ['line', 'Line'],
  ['area', 'Area'],
  ['pie', 'Pie'],
  ['doughnut', 'Doughnut'],
  ['scatter', 'Scatter'],
  ['bubble', 'Bubble'],
  ['radar', 'Radar'],
  ['filledRadar', 'Filled Radar'],
];

const AGGREGATE_LABELS: readonly (readonly [ChartAggregate, string])[] = [
  ['sum', 'Sum'],
  ['average', 'Average'],
  ['count', 'Count'],
  ['min', 'Min'],
  ['max', 'Max'],
];

const STACKING_LABELS: readonly (readonly [ChartStacking, string])[] = [
  ['none', 'None'],
  ['stacked', 'Stacked'],
  ['percent', '100% stacked'],
];

/** The props binding a text input to a draft field (preview on input, commit on change/Enter). */
interface TextInputProps {
  readonly type: 'text';
  readonly value: string;
  readonly onInput: (e: Event) => void;
  readonly onChange: () => void;
  readonly onKeyDown: (e: KeyboardEvent) => void;
}

const blank = (text: string): string | null => (text.trim() === '' ? null : text.trim());

/** A titled field with a hover description on both the title and the control. */
function Field(props: { label: string; tip: string; children: ComponentChildren }) {
  return (
    <label class="cs-chart-field" title={props.tip}>
      <span>{props.label}</span>
      {props.children}
    </label>
  );
}

/**
 * The chart sheet page: the chart fills the left, and a panel on the right edits its data and
 * presentation. Title and type are pinned at the top; the rest adapts to the type's family so each
 * chart only asks for the data it plots. A Table / Range toggle picks how the data is given: by a
 * named table's columns (`ChartTableFields`), or by cell ranges typed or picked on the sheet. Text
 * fields preview while typing and commit on change (blur/Enter); pickers, the legend toggle, and
 * add/remove series commit at once.
 */
export function ChartSheetView(props: ChartSheetViewProps) {
  const source = props.sheet.chart;
  const [draft, setDraft] = useState<ChartDefinition>(() => source ?? createChart());
  const draftRef = useRef(draft);
  useEffect(() => {
    if (source && source !== draftRef.current) {
      draftRef.current = source;
      setDraft(source);
    }
  }, [source]);
  // Table or Range: a bound chart is always in table mode; otherwise the user's last choice, else
  // table mode when there is a table to plot and nothing plotted yet.
  const [chosenMode, setChosenMode] = useState<'table' | 'range' | null>(null);
  const [modeError, setModeError] = useState<string | null>(null);
  useEffect(() => {
    setChosenMode(null);
    setModeError(null);
  }, [props.sheet]);
  const tableMode =
    draft.source !== null ||
    (chosenMode ?? (ChartTableBinding.startsInTableMode(props.workbook, draft) ? 'table' : 'range')) ===
      'table';
  /** Replace the whole chart and commit it (table-mode picks, the mode toggle, the type). */
  const replaceChart = (next: ChartDefinition) => {
    draftRef.current = next;
    setDraft(next);
    setModeError(null);
    props.onChange(next);
  };
  const switchMode = (mode: 'table' | 'range') => {
    if (mode === 'range') {
      setChosenMode('range');
      if (draft.source) replaceChart(ChartTableBinding.toRanges(props.workbook, draft));
      return;
    }
    if (tableMode) return;
    const result = ChartTableBinding.toTable(props.workbook, draft);
    if (result.ok) {
      setChosenMode('table');
      replaceChart(result.chart);
    } else {
      setModeError(result.reason);
    }
  };
  /** A "Pick on sheet" button for `field`, when the host supports picking. */
  const pick = (field: ChartField) =>
    props.onPick && (
      <button
        type="button"
        class="cs-btn cs-chart-pick"
        title="Pick on sheet: select the cells on the data sheet, then press Enter. Esc cancels."
        aria-label="Pick on sheet"
        onClick={() => props.onPick!(field)}
      >
        ⌖
      </button>
    );

  const update = (mutate: (c: ChartDefinition) => void, commit: boolean) => {
    const next = cloneChart(draftRef.current);
    mutate(next);
    draftRef.current = next;
    setDraft(next);
    if (commit) props.onChange(next);
  };
  const commitDraft = () => {
    if (draftRef.current !== source) props.onChange(draftRef.current);
  };
  /** Props for a text input bound to a draft field. */
  const text = (value: string, set: (c: ChartDefinition, v: string) => void): TextInputProps => ({
    type: 'text',
    value,
    onInput: (e: Event) => {
      const v = (e.currentTarget as HTMLInputElement).value;
      update((c) => set(c, v), false);
    },
    onChange: commitDraft,
    onKeyDown: (e: KeyboardEvent) => {
      if (e.key === 'Enter') commitDraft();
    },
  });

  // Re-read cells only when what the data depends on changes, not on every title keystroke.
  const dataKey = JSON.stringify([
    draft.type,
    draft.categoriesReference,
    draft.series,
    draft.source,
    draft.aggregate,
  ]);
  const data = useMemo(
    () => ChartDataResolver.resolve(props.workbook, draft),
    // `version` stands in for workbook content changes.
    [props.workbook, dataKey, props.version],
  );
  const resolved = useMemo(
    () => ({
      ...data,
      title: draft.title,
      stacking: draft.stacking,
      showLegend: draft.showLegend,
      xAxisTitle: draft.xAxisTitle,
      yAxisTitle: draft.yAxisTitle,
    }),
    [data, draft],
  );

  const host = useRef<HTMLDivElement>(null);
  const [size, setSize] = useState({ width: 0, height: 0 });
  useLayoutEffect(() => {
    const el = host.current;
    if (!el) return;
    const measure = () => setSize({ width: el.clientWidth, height: el.clientHeight });
    measure();
    if (typeof ResizeObserver === 'undefined') return;
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const family = chartFamily(draft.type);
  const categoryField = (
    <Field {...categoryCopy(family)}>
      <div class="cs-chart-ref">
        <input
          placeholder="e.g. Sheet1.A2:A7"
          {...text(draft.categoriesReference ?? '', (c, v) => {
            c.categoriesReference = blank(v);
          })}
        />
        {pick({ kind: 'categories' })}
      </div>
    </Field>
  );

  return (
    <div class="cs-chart-sheet">
      <div class="cs-chart-host" ref={host}>
        {size.width > 0 && size.height > 0 && (
          <ChartSvg chart={resolved} width={size.width} height={size.height} dark={props.dark} interactive />
        )}
      </div>
      <aside class="cs-chart-panel" aria-label="Chart data">
        <h2>Chart data</h2>
        <div class="cs-chart-mode" role="radiogroup" aria-label="Plot from">
          {(['table', 'range'] as const).map((mode) => (
            <button
              key={mode}
              type="button"
              role="radio"
              aria-checked={tableMode === (mode === 'table')}
              class={tableMode === (mode === 'table') ? 'cs-chart-mode-on' : ''}
              title={
                mode === 'table'
                  ? 'Plot a named table by picking its columns. The chart follows the table as it grows.'
                  : 'Plot cell ranges you type or pick on the sheet.'
              }
              onClick={() => switchMode(mode)}
            >
              {mode === 'table' ? 'Table' : 'Range'}
            </button>
          ))}
        </div>
        {modeError && (
          <p class="cs-chart-hint cs-chart-error" role="alert">
            {modeError}
          </p>
        )}
        <Field label="Title" tip="The heading shown above the chart. Leave blank for no title.">
          <input
            {...text(draft.title, (c, v) => {
              c.title = v;
            })}
          />
        </Field>
        <Field
          label="Chart type"
          tip={
            'How the data is drawn. Column/Bar are the same data shown vertically or horizontally; ' +
            'Line/Area trace it across categories; Pie/Doughnut show one series as proportions; ' +
            'Scatter/Bubble plot X vs Y; Radar plots each category on its own spoke. ' +
            'The fields below change to match the type you pick.'
          }
        >
          <select
            value={draft.type}
            onChange={(e) => {
              const v = (e.currentTarget as HTMLSelectElement).value as ChartType;
              replaceChart(ChartTableBinding.retype(props.workbook, draftRef.current, v));
            }}
          >
            {CHART_TYPE_LABELS.map(([t, label]) => (
              <option key={t} value={t}>
                {label}
              </option>
            ))}
          </select>
        </Field>

        {ChartTypes.supportsStacking(draft.type) && (
          <Field
            label="Stacking"
            tip="Stack each category's series on top of one another, or scale every category to 100%."
          >
            <select
              value={draft.stacking}
              onChange={(e) => {
                const v = (e.currentTarget as HTMLSelectElement).value as ChartStacking;
                update((c) => {
                  c.stacking = v;
                }, true);
              }}
            >
              {STACKING_LABELS.map(([s, label]) => (
                <option key={s} value={s}>
                  {label}
                </option>
              ))}
            </select>
          </Field>
        )}

        {ChartTypes.groupsCategories(draft.type) && (
          <Field
            label="Aggregate"
            tip="Rows with the same category label become one category. This is how their values combine."
          >
            <select
              value={draft.aggregate}
              onChange={(e) => {
                const v = (e.currentTarget as HTMLSelectElement).value as ChartAggregate;
                update((c) => {
                  c.aggregate = v;
                }, true);
              }}
            >
              {AGGREGATE_LABELS.map(([a, label]) => (
                <option key={a} value={a}>
                  {label}
                </option>
              ))}
            </select>
          </Field>
        )}

        {(family === 'cartesian' || family === 'xy') && (
          <>
            <Field label="X-axis title" tip="Optional label drawn beneath the horizontal axis.">
              <input
                {...text(draft.xAxisTitle ?? '', (c, v) => {
                  c.xAxisTitle = blank(v);
                })}
              />
            </Field>
            <Field label="Y-axis title" tip="Optional label for the value axis.">
              <input
                {...text(draft.yAxisTitle ?? '', (c, v) => {
                  c.yAxisTitle = blank(v);
                })}
              />
            </Field>
          </>
        )}

        <label
          class="cs-chart-check"
          title={
            family === 'partToWhole'
              ? 'Show a key naming each slice.'
              : 'Show a key naming each series. Hide it when a single series needs no label.'
          }
        >
          <input
            type="checkbox"
            checked={draft.showLegend}
            onChange={(e) => {
              const v = (e.currentTarget as HTMLInputElement).checked;
              update((c) => {
                c.showLegend = v;
              }, true);
            }}
          />
          Show legend
        </label>

        <hr />

        {tableMode ? (
          <ChartTableFields
            workbook={props.workbook}
            version={props.version}
            chart={draft}
            family={family}
            onChange={replaceChart}
          />
        ) : family === 'partToWhole' ? (
          <>
            {categoryField}
            <h3>Values</h3>
            <p class="cs-chart-hint">A pie or doughnut shows one set of values as proportions of a whole.</p>
            <Field
              label="Slice values"
              tip="The cells holding the value for each slice — one per slice label, e.g. Sheet1.B2:B7."
            >
              <div class="cs-chart-ref">
                <input
                  placeholder="e.g. Sheet1.B2:B7"
                  {...text(draft.series[0]?.valuesReference ?? '', (c, v) => {
                    // A part-to-whole chart only draws its first series; create it on first edit.
                    const first = c.series[0];
                    if (first) first.valuesReference = v;
                    else c.series.push(createSeries({ valuesReference: v }));
                  })}
                />
                {pick({ kind: 'values', series: 0 })}
              </div>
            </Field>
          </>
        ) : (
          <>
            {family !== 'xy' && categoryField}
            <h3>Series</h3>
            <p class="cs-chart-hint">
              {family === 'xy'
                ? 'Each series is a set of points. Give each its own X and Y values.'
                : family === 'radar'
                  ? 'Each series is one outline plotted across the spokes.'
                  : 'Each series is one set of values plotted across the categories.'}
            </p>
            <div class="cs-chart-series-list">
              {draft.series.map((series, i) => (
                <SeriesRow
                  key={i}
                  series={series}
                  ordinal={i + 1}
                  family={family}
                  bubble={draft.type === 'bubble'}
                  chartType={draft.type}
                  text={(value, set) => text(value, (c, v) => set(c.series[i]!, v))}
                  pick={(kind) => pick({ kind, series: i })}
                  onTypeChange={(t) =>
                    update((c) => {
                      c.series[i]!.type = t;
                    }, true)
                  }
                  onRemove={() =>
                    update((c) => {
                      c.series.splice(i, 1);
                    }, true)
                  }
                />
              ))}
            </div>
            <button
              type="button"
              class="cs-btn cs-chart-add"
              title={
                family === 'xy'
                  ? 'Add another set of points to plot. Give each its own X and Y values.'
                  : 'Add another set of values to plot (a second column, line, etc.). Give each its own values range.'
              }
              onClick={() =>
                update((c) => {
                  c.series.push(createSeries({ valuesReference: '' }));
                }, true)
              }
            >
              + Add series
            </button>
          </>
        )}
      </aside>
    </div>
  );
}

function categoryCopy(family: ChartFamily): { label: string; tip: string } {
  switch (family) {
    case 'partToWhole':
      return {
        label: 'Slice labels',
        tip: 'The cells naming each slice — one per value, e.g. Sheet1.A2:A7. Shown in the legend.',
      };
    case 'radar':
      return {
        label: 'Spoke labels',
        tip: 'The cells naming each spoke of the radar — one per value, e.g. Sheet1.A2:A7.',
      };
    case 'cartesian':
    case 'xy':
    default:
      return {
        label: 'Category labels',
        tip:
          'The cells holding the label for each data point — one per row, e.g. Sheet1.A2:A7. ' +
          'These name the columns/bars/points along the category axis.',
      };
  }
}

/**
 * The types one series of a column/bar/line/area chart can be drawn as: the chart's bar
 * orientation, a line or an area. Pie, scatter and the rest can't share category axes, so they
 * aren't offered (one already set, e.g. by a file, stays listed so the choice shows).
 */
export function drawAsOptions(
  chartType: ChartType,
  current: ChartType | null,
): (readonly [ChartType, string])[] {
  const bar: ChartType = chartType === 'bar' ? 'bar' : 'column';
  const allowed = new Set<ChartType>([bar, 'line', 'area']);
  if (current) allowed.add(current);
  return CHART_TYPE_LABELS.filter(([t]) => allowed.has(t));
}

type TextBinder = (value: string, set: (s: ChartSeries, v: string) => void) => TextInputProps;

function SeriesRow(props: {
  series: ChartSeries;
  ordinal: number;
  family: ChartFamily;
  bubble: boolean;
  chartType: ChartType;
  text: TextBinder;
  /** The "Pick on sheet" button for one of this series' reference fields. */
  pick: (kind: ChartSeriesFieldKind) => ComponentChildren;
  onTypeChange: (type: ChartType | null) => void;
  onRemove: () => void;
}) {
  const { series, family } = props;
  return (
    <fieldset class="cs-chart-series">
      <legend>
        <span>Series {props.ordinal}</span>
        <button
          type="button"
          class="cs-btn danger"
          title="Remove this series from the chart."
          aria-label={`Remove series ${props.ordinal}`}
          onClick={props.onRemove}
        >
          ✕
        </button>
      </legend>
      <Field
        label="Name"
        tip={'A label for this series. Leave blank to use the name cell below, or an automatic "Series N".'}
      >
        <input
          placeholder="Series name (optional)"
          {...props.text(series.name ?? '', (s, v) => {
            s.name = blank(v);
          })}
        />
      </Field>
      <Field
        label="Name cell"
        tip="A single cell whose text names the series (e.g. the column header Sheet1.B1). Overrides the typed name."
      >
        <div class="cs-chart-ref">
          <input
            placeholder="e.g. Sheet1.B1"
            {...props.text(series.nameReference ?? '', (s, v) => {
              s.nameReference = blank(v);
            })}
          />
          {props.pick('nameCell')}
        </div>
      </Field>
      {family === 'xy' && (
        <Field
          label="X values"
          tip="The horizontal position of each point, e.g. Sheet1.A2:A7. Defaults to 1, 2, 3… when blank."
        >
          <div class="cs-chart-ref">
            <input
              placeholder="e.g. Sheet1.A2:A7"
              {...props.text(series.xValuesReference ?? '', (s, v) => {
                s.xValuesReference = blank(v);
              })}
            />
            {props.pick('xValues')}
          </div>
        </Field>
      )}
      <Field
        label={family === 'xy' ? 'Y values' : 'Values'}
        tip={
          family === 'xy'
            ? 'The vertical position of each point — one per X value, e.g. Sheet1.B2:B7.'
            : "The cells holding this series' numbers — one per category, e.g. Sheet1.B2:B7. This is the data that gets plotted."
        }
      >
        <div class="cs-chart-ref">
          <input
            placeholder="e.g. Sheet1.B2:B7"
            {...props.text(series.valuesReference, (s, v) => {
              s.valuesReference = v;
            })}
          />
          {props.pick('values')}
        </div>
      </Field>
      {family === 'cartesian' && (
        <Field
          label="Draw as"
          tip="Override the chart type for just this series — e.g. draw one series as a Line over Column bars to make a combo chart."
        >
          <select
            value={series.type ?? ''}
            onChange={(e) => {
              const v = (e.currentTarget as HTMLSelectElement).value;
              props.onTypeChange(v === '' ? null : (v as ChartType));
            }}
          >
            <option value="">Same as chart</option>
            {drawAsOptions(props.chartType, series.type ?? null).map(([t, label]) => (
              <option key={t} value={t}>
                {label}
              </option>
            ))}
          </select>
        </Field>
      )}
      {props.bubble && (
        <Field label="Bubble size" tip="The value that sizes each bubble — one per point, e.g. Sheet1.C2:C7.">
          <div class="cs-chart-ref">
            <input
              placeholder="e.g. Sheet1.C2:C7"
              {...props.text(series.bubbleSizeReference ?? '', (s, v) => {
                s.bubbleSizeReference = blank(v);
              })}
            />
            {props.pick('bubbleSizes')}
          </div>
        </Field>
      )}
    </fieldset>
  );
}
