import type { ChartDefinition, ChartSeries, ChartStacking, ChartType } from '@cellestra/charts';
import { cloneChart, createChart, createSeries, ChartTypes } from '@cellestra/charts';
import { ChartDataResolver, type Workbook, type Worksheet } from '@cellestra/core';
import type { ComponentChildren } from 'preact';
import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'preact/hooks';
import { ChartSvg } from './ChartSvg.tsx';

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
 * chart only asks for the data it plots. Text fields preview while typing and commit on change
 * (blur/Enter); pickers, the legend toggle, and add/remove series commit at once.
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
  const dataKey = JSON.stringify([draft.type, draft.categoriesReference, draft.series]);
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

  return (
    <div class="cs-chart-sheet">
      <div class="cs-chart-host" ref={host}>
        {size.width > 0 && size.height > 0 && (
          <ChartSvg chart={resolved} width={size.width} height={size.height} dark={props.dark} />
        )}
      </div>
      <aside class="cs-chart-panel" aria-label="Chart data">
        <h2>Chart data</h2>
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
              update((c) => {
                c.type = v;
              }, true);
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

        {family !== 'xy' && (
          <Field {...categoryCopy(family)}>
            <input
              placeholder="e.g. Sheet1.A2:A7"
              {...text(draft.categoriesReference ?? '', (c, v) => {
                c.categoriesReference = blank(v);
              })}
            />
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

        {family === 'partToWhole' ? (
          <>
            <h3>Values</h3>
            <p class="cs-chart-hint">A pie or doughnut shows one set of values as proportions of a whole.</p>
            <Field
              label="Slice values"
              tip="The cells holding the value for each slice — one per slice label, e.g. Sheet1.B2:B7."
            >
              <input
                placeholder="e.g. Sheet1.B2:B7"
                {...text(draft.series[0]?.valuesReference ?? '', (c, v) => {
                  // A part-to-whole chart only draws its first series; create it on first edit.
                  const first = c.series[0];
                  if (first) first.valuesReference = v;
                  else c.series.push(createSeries({ valuesReference: v }));
                })}
              />
            </Field>
          </>
        ) : (
          <>
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
                  text={(value, set) => text(value, (c, v) => set(c.series[i]!, v))}
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
    default:
      return {
        label: 'Category labels',
        tip:
          'The cells holding the label for each data point — one per row, e.g. Sheet1.A2:A7. ' +
          'These name the columns/bars/points along the category axis.',
      };
  }
}

type TextBinder = (value: string, set: (s: ChartSeries, v: string) => void) => TextInputProps;

function SeriesRow(props: {
  series: ChartSeries;
  ordinal: number;
  family: ChartFamily;
  bubble: boolean;
  text: TextBinder;
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
        <input
          placeholder="e.g. Sheet1.B1"
          {...props.text(series.nameReference ?? '', (s, v) => {
            s.nameReference = blank(v);
          })}
        />
      </Field>
      {family === 'xy' && (
        <Field
          label="X values"
          tip="The horizontal position of each point, e.g. Sheet1.A2:A7. Defaults to 1, 2, 3… when blank."
        >
          <input
            placeholder="e.g. Sheet1.A2:A7"
            {...props.text(series.xValuesReference ?? '', (s, v) => {
              s.xValuesReference = blank(v);
            })}
          />
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
        <input
          placeholder="e.g. Sheet1.B2:B7"
          {...props.text(series.valuesReference, (s, v) => {
            s.valuesReference = v;
          })}
        />
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
            {CHART_TYPE_LABELS.map(([t, label]) => (
              <option key={t} value={t}>
                {label}
              </option>
            ))}
          </select>
        </Field>
      )}
      {props.bubble && (
        <Field label="Bubble size" tip="The value that sizes each bubble — one per point, e.g. Sheet1.C2:C7.">
          <input
            placeholder="e.g. Sheet1.C2:C7"
            {...props.text(series.bubbleSizeReference ?? '', (s, v) => {
              s.bubbleSizeReference = blank(v);
            })}
          />
        </Field>
      )}
    </fieldset>
  );
}
