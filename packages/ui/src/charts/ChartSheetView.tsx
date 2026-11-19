import type {
  ChartAggregate,
  ChartDefinition,
  ChartSeries,
  ChartStacking,
  ChartType,
} from '@cellestra/charts';
import { cloneChart, createChart, createSeries, ChartTypes } from '@cellestra/charts';
import {
  ChartChecks,
  ChartDataResolver,
  ChartFields,
  ChartTableBinding,
  quoteSheetIfNeeded,
  TableCommands,
  tableNameError,
  type CellRange,
  type ChartField,
  type ChartSeriesFieldKind,
  type FromRangesResult,
  type Workbook,
  type Worksheet,
} from '@cellestra/core';
import type { ComponentChildren } from 'preact';
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'preact/hooks';
import { Icon } from '../components/Icon.tsx';
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
  /**
   * A committed edit: the host applies it (e.g. via `ChartEditCommands.setChart`). `notice` is set
   * for an edit worth announcing with Undo, such as removing a series ("Removed Sales").
   */
  readonly onChange: (next: ChartDefinition, notice?: string) => void;
  /**
   * "Pick on sheet" for a reference field: the host shows the data sheet, lets the user select
   * cells, and sets the field. Omitted, the buttons don't show.
   */
  readonly onPick?: (field: ChartField) => void;
  /**
   * "Make a table": the host makes the chart's data a table called `name` and binds the chart to it
   * (`ChartEditCommands.makeTableForChart`), returning why it couldn't, or null. Omitted, the panel
   * doesn't offer it.
   */
  readonly onMakeTable?: (name: string) => string | null;
  /**
   * "Edit Data Range…": the host shows the data sheet, lets the user select one block, and re-plots
   * the chart from it. Omitted, the button doesn't show.
   */
  readonly onEditDataRange?: () => void;
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

/** The icon drawn for each chart type in the type picker. */
export const CHART_TYPE_ICONS: Record<ChartType, string> = {
  column: 'chart-column',
  bar: 'chart-bar',
  line: 'chart-line',
  area: 'chart-area',
  pie: 'chart-pie',
  doughnut: 'chart-doughnut',
  scatter: 'chart-scatter',
  bubble: 'chart-bubble',
  radar: 'chart-radar',
  filledRadar: 'chart-filled-radar',
};

const AGGREGATE_LABELS: readonly (readonly [ChartAggregate, string])[] = [
  ['sum', 'Add up'],
  ['average', 'Average'],
  ['count', 'Count'],
  ['min', 'Lowest'],
  ['max', 'Highest'],
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

/** A stable key for a reference field, for matching problems and naming their messages. */
export function fieldKey(field: ChartField): string {
  return field.kind === 'categories' ? 'categories' : `${field.kind}-${field.series}`;
}

/** A region as a reference, like `Data.A1:C7`. */
const where = (region: { readonly sheet: string; readonly range: CellRange }): string =>
  `${quoteSheetIfNeeded(region.sheet)}.${region.range.toString()}`;

const problemId = (field: ChartField): string => `cs-chart-problem-${fieldKey(field)}`;

/**
 * A titled field with a hover description on both the title and the control. A problem shows
 * beneath it, and the control (given `invalidProps`) points at it for screen readers.
 */
function Field(props: {
  label: string;
  tip: string;
  problem?: { readonly id: string; readonly text: string } | null;
  /** What the field's reference points at ("Jan, Feb, Mar … 12 cells"), shown when it has no problem. */
  sample?: string | null;
  children: ComponentChildren;
}) {
  return (
    <div class="cs-chart-field-wrap">
      <label class="cs-chart-field" title={props.tip}>
        <span>{props.label}</span>
        {props.children}
      </label>
      {props.problem ? (
        <p class="cs-chart-problem" id={props.problem.id}>
          {props.problem.text}
        </p>
      ) : (
        props.sample && <p class="cs-chart-sample">{props.sample}</p>
      )}
    </div>
  );
}

/** The chart types as a row of pictures, one radio button each. */
function ChartTypePicker(props: {
  value: ChartType;
  /** Types the data can't fill, with why. */
  unavailable: ReadonlyMap<ChartType, string>;
  onPick: (type: ChartType) => void;
}) {
  const buttons = useRef<(HTMLButtonElement | null)[]>([]);
  // Arrow keys move focus only; Space, Enter or a click picks. Picking on every arrow press would
  // add an undo step per type passed and let a pie on the way drop a bound chart's columns.
  const [focused, setFocused] = useState<ChartType | null>(null);
  const usable = CHART_TYPE_LABELS.map(([t]) => t).filter(
    (t) => t === props.value || !props.unavailable.has(t),
  );
  const current = focused !== null && usable.includes(focused) ? focused : props.value;
  const step = (by: number | 'first' | 'last') => {
    const at = usable.indexOf(current);
    const i =
      by === 'first' ? 0 : by === 'last' ? usable.length - 1 : (at + by + usable.length) % usable.length;
    const next = usable[i]!;
    setFocused(next);
    buttons.current[CHART_TYPE_LABELS.findIndex(([t]) => t === next)]?.focus();
  };
  const keys: Record<string, number | 'first' | 'last'> = {
    ArrowRight: 1,
    ArrowDown: 1,
    ArrowLeft: -1,
    ArrowUp: -1,
    Home: 'first',
    End: 'last',
  };
  return (
    <div class="cs-chart-field-wrap">
      <span class="cs-chart-label" id="cs-chart-type-label">
        Chart type
      </span>
      <div class="cs-chart-types" role="radiogroup" aria-labelledby="cs-chart-type-label">
        {CHART_TYPE_LABELS.map(([type, label], i) => {
          const on = type === props.value;
          const why = on ? undefined : props.unavailable.get(type);
          return (
            <button
              key={type}
              ref={(el) => {
                buttons.current[i] = el;
              }}
              type="button"
              role="radio"
              aria-checked={on}
              aria-disabled={why ? true : undefined}
              aria-label={label}
              tabIndex={type === current ? 0 : -1}
              title={why ? `${label}: ${why}` : label}
              class={on ? 'cs-chart-type cs-chart-type-on' : 'cs-chart-type'}
              onClick={() => {
                setFocused(null);
                if (!why && !on) props.onPick(type);
              }}
              onKeyDown={(e) => {
                const by = keys[e.key];
                if (by === undefined) return;
                e.preventDefault();
                step(by);
              }}
            >
              <Icon name={CHART_TYPE_ICONS[type]} />
              <span>{label}</span>
            </button>
          );
        })}
      </div>
    </div>
  );
}

/**
 * "Make a table": the offer to turn the chart's cells into a named table, so the chart plots
 * columns by name and follows the table as it grows. Opened, it asks for the table's name, filled
 * in from the labels above the data.
 */
function MakeTableOffer(props: {
  workbook: Workbook;
  region: { readonly sheet: string; readonly range: CellRange; readonly header: readonly string[] | null };
  open: boolean;
  onOpen: () => void;
  onClose: () => void;
  onCreate: (name: string) => string | null;
}) {
  const { region } = props;
  const [name, setName] = useState(() =>
    TableCommands.suggestName(props.workbook, region.sheet, region.range),
  );
  const [failure, setFailure] = useState<string | null>(null);
  // The name field only exists while the form is open, so it takes focus as it appears.
  const focusOnOpen = useCallback((el: HTMLInputElement | null) => {
    el?.focus();
    el?.select();
  }, []);
  const place = where(region);
  if (!props.open)
    return (
      <div class="cs-chart-offer">
        <p>
          Plot by column names? Make {place} a table, and the chart follows its columns as rows are added.
        </p>
        <button type="button" class="cs-btn" onClick={props.onOpen}>
          Make a Table…
        </button>
      </div>
    );
  const error =
    name.trim() === '' ? 'Give the table a name.' : (tableNameError(props.workbook, name.trim()) ?? failure);
  const create = () => {
    if (error) return;
    setFailure(props.onCreate(name.trim()));
  };
  return (
    <form
      class="cs-chart-offer"
      aria-label="Make a table"
      onSubmit={(e) => {
        e.preventDefault();
        create();
      }}
    >
      <p>Make {place} a table. The chart then plots its columns by name.</p>
      <p class="cs-chart-sample">
        {region.header
          ? `Column names: ${region.header.map((h, i) => h || `Column${i + 1}`).join(', ')}.`
          : 'The data has no row of labels, so a header row named Column 1, Column 2… is added above it.'}
      </p>
      <div class="cs-chart-field-wrap">
        <label class="cs-chart-field">
          <span>Table name</span>
          <input
            ref={focusOnOpen}
            type="text"
            value={name}
            aria-invalid={error ? true : undefined}
            aria-describedby={error ? 'cs-chart-table-name-problem' : undefined}
            onInput={(e) => {
              setName((e.currentTarget as HTMLInputElement).value);
              setFailure(null);
            }}
            onKeyDown={(e) => {
              if (e.key !== 'Escape' || e.isComposing) return;
              e.preventDefault();
              e.stopPropagation();
              props.onClose();
            }}
          />
        </label>
        {error && (
          <p class="cs-chart-problem" id="cs-chart-table-name-problem">
            {error}
          </p>
        )}
      </div>
      <div class="cs-chart-offer-buttons">
        <button type="submit" class="cs-btn primary" disabled={error !== null}>
          Create Table
        </button>
        <button type="button" class="cs-btn" onClick={props.onClose}>
          Cancel
        </button>
      </div>
    </form>
  );
}

/**
 * The chart sheet page: the chart fills the left, and a panel on the right edits its data and
 * presentation. A line at the top says where the data comes from; then the title, the type (as
 * pictures) and the data, which adapts to the type's family so each chart only asks for what it
 * plots. A Table / Range toggle picks how the data is given: by a named table's columns
 * (`ChartTableFields`), or by cell ranges typed or picked on the sheet. Ranges that can't plot say
 * why under their field, and cells that could be a table get a "Make a Table" offer. Text fields
 * preview while typing and commit on change (blur/Enter); pickers, toggles, and add/remove series
 * commit at once.
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
  // Why the last switch to Table failed, kept only while the chart and cells are as they were then.
  const [modeFailure, setModeFailure] = useState<{
    readonly result: Extract<FromRangesResult, { ok: false }>;
    readonly key: string;
  } | null>(null);
  const [offerOpen, setOfferOpen] = useState(false);
  useEffect(() => {
    setChosenMode(null);
    setModeFailure(null);
    setOfferOpen(false);
  }, [props.sheet]);
  const tableMode =
    draft.source !== null ||
    (chosenMode ?? (ChartTableBinding.startsInTableMode(props.workbook, draft) ? 'table' : 'range')) ===
      'table';
  /** Replace the whole chart and commit it (table-mode picks, the mode toggle, the type). */
  const replaceChart = (next: ChartDefinition) => {
    draftRef.current = next;
    setDraft(next);
    setModeFailure(null);
    props.onChange(next);
  };

  // Re-read cells only when what the data depends on changes, not on every title keystroke.
  const dataKey = JSON.stringify([
    draft.type,
    draft.categoriesReference,
    draft.series,
    draft.source,
    draft.aggregate,
  ]);
  // A key per series row, so a row's open/folded state stays with its series when one above is
  // removed. Removing and adding keep the keys in step; a change from elsewhere that adds or drops
  // series (undo, Swap Rows and Columns) starts the rows afresh.
  const rowKeys = useRef<number[]>([]);
  const nextRowKey = useRef(0);
  const newRowKey = () => nextRowKey.current++;
  if (rowKeys.current.length !== draft.series.length) rowKeys.current = draft.series.map(newRowKey);

  const failureKey = `${dataKey}|${props.version}`;
  const failure = modeFailure?.key === failureKey ? modeFailure.result : null;
  const data = useMemo(
    () => ChartDataResolver.resolve(props.workbook, draft),
    [props.workbook, dataKey, props.version],
  );
  const problems = useMemo(() => {
    const byField = new Map<string, string>();
    for (const p of ChartChecks.problems(props.workbook, draft))
      if (!byField.has(fieldKey(p.field))) byField.set(fieldKey(p.field), p.message);
    return byField;
  }, [props.workbook, dataKey, props.version]);
  const summary = useMemo(
    () => ChartChecks.summary(props.workbook, draft),
    [props.workbook, dataKey, props.version],
  );
  const region = useMemo(
    () => (props.onMakeTable ? ChartTableBinding.tableRegion(props.workbook, draft) : null),
    [props.workbook, dataKey, props.version, props.onMakeTable !== undefined],
  );
  const unavailable = useMemo(
    () => ChartTableBinding.unavailableTypes(props.workbook, draft),
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

  const switchMode = (mode: 'table' | 'range') => {
    setOfferOpen(false);
    if (mode === 'range') {
      setChosenMode('range');
      setModeFailure(null);
      if (draft.source) replaceChart(ChartTableBinding.toRanges(props.workbook, draft));
      return;
    }
    if (tableMode) return;
    const result = ChartTableBinding.toTable(props.workbook, draft);
    if (result.ok) {
      setChosenMode('table');
      replaceChart(result.chart);
    } else if ((result.code === 'notInTable' || result.code === 'noTable') && region) {
      // The data isn't a table yet: ask for a name and make it one.
      setModeFailure(null);
      setOfferOpen(true);
    } else {
      setModeFailure({ result, key: failureKey });
    }
  };
  /** "Use Whole Columns": bind each range to the table column it sits in. */
  const bindWholeColumns = () => {
    const result = ChartTableBinding.fromRanges(props.workbook, draft, { wholeColumns: true });
    if (result.ok) {
      setChosenMode('table');
      replaceChart(result.chart);
    } else setModeFailure({ result, key: failureKey });
  };
  const makeTable = (name: string): string | null => {
    const why = props.onMakeTable?.(name) ?? null;
    if (why === null) {
      setOfferOpen(false);
      setChosenMode('table');
    }
    return why;
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
  const problemOf = (field: ChartField) => {
    const text = problems.get(fieldKey(field));
    return text ? { id: problemId(field), text } : null;
  };
  /** The first cells a field's reference points at, re-read as it is typed. */
  const sampleOf = (field: ChartField) => ChartChecks.sample(props.workbook, ChartFields.get(draft, field));
  /** `aria-invalid` and the message's id for a reference input with a problem. */
  const invalid = (field: ChartField) =>
    problems.has(fieldKey(field))
      ? { 'aria-invalid': true as const, 'aria-describedby': problemId(field) }
      : {};

  /** The draft as it is committed: whole columns typed as `A:A` written out as cells. */
  const settle = (draftNow: ChartDefinition): ChartDefinition => {
    const next = ChartFields.expandWholeLines(draftNow);
    if (next !== draftNow) {
      draftRef.current = next;
      setDraft(next);
    }
    return next;
  };
  const update = (mutate: (c: ChartDefinition) => void, commit: boolean, notice?: string) => {
    const next = cloneChart(draftRef.current);
    mutate(next);
    draftRef.current = next;
    setDraft(next);
    if (commit) props.onChange(settle(next), notice);
  };
  const commitDraft = () => {
    const next = settle(draftRef.current);
    if (next !== source) props.onChange(next);
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
    <Field
      {...categoryCopy(family)}
      problem={problemOf({ kind: 'categories' })}
      sample={sampleOf({ kind: 'categories' })}
    >
      <div class="cs-chart-ref">
        <input
          placeholder="e.g. Sheet1.A2:A7"
          {...invalid({ kind: 'categories' })}
          {...text(draft.categoriesReference ?? '', (c, v) => {
            c.categoriesReference = blank(v);
          })}
        />
        {pick({ kind: 'categories' })}
      </div>
    </Field>
  );
  const sliceValues = { kind: 'values', series: 0 } as const;

  return (
    <div class="cs-chart-sheet">
      <div class="cs-chart-host" ref={host}>
        {size.width > 0 && size.height > 0 && (
          <ChartSvg chart={resolved} width={size.width} height={size.height} dark={props.dark} interactive />
        )}
      </div>
      <aside class="cs-chart-panel" aria-label="Chart data">
        <h2>Chart data</h2>
        <p class="cs-chart-summary" role="status">
          {summary}
        </p>
        {data.series.length === 0 && (
          <p class="cs-chart-hint cs-chart-howto">
            {tableMode
              ? 'Pick a table below, then tick the columns to plot.'
              : 'Type a range such as Sheet1.B2:B7 into a field and press Enter, or press ⌖ beside it to select the cells on the sheet.'}
          </p>
        )}
        <Field label="Title" tip="The heading shown above the chart. Leave blank for no title.">
          <input
            {...text(draft.title, (c, v) => {
              c.title = v;
            })}
          />
        </Field>
        <ChartTypePicker
          value={draft.type}
          unavailable={unavailable}
          onPick={(type) => replaceChart(ChartTableBinding.retype(props.workbook, draftRef.current, type))}
        />

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

        <hr />

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
        {failure && (
          <div class="cs-chart-offer cs-chart-offer-error" role="alert">
            <p>{failure.reason}</p>
            {failure.code === 'notWholeColumns' && (
              <button type="button" class="cs-btn" onClick={bindWholeColumns}>
                Use Whole Columns
              </button>
            )}
          </div>
        )}
        {props.onEditDataRange && (
          <button
            type="button"
            class="cs-btn cs-chart-data-range"
            title="Select one block on the sheet: labels in the first column, a series in each other column."
            onClick={props.onEditDataRange}
          >
            Edit Data Range…
          </button>
        )}
        {!tableMode && region && props.onMakeTable && (
          <MakeTableOffer
            key={where(region)}
            workbook={props.workbook}
            region={region}
            open={offerOpen}
            onOpen={() => setOfferOpen(true)}
            onClose={() => setOfferOpen(false)}
            onCreate={makeTable}
          />
        )}

        {tableMode ? (
          <ChartTableFields
            workbook={props.workbook}
            version={props.version}
            chart={draft}
            family={family}
            onChange={replaceChart}
            onKeepCells={() => switchMode('range')}
          />
        ) : family === 'partToWhole' ? (
          <>
            {categoryField}
            <Field
              label="Slice values"
              tip="The cells holding the value for each slice: one per slice label, e.g. Sheet1.B2:B7."
              problem={problemOf(sliceValues)}
              sample={sampleOf(sliceValues)}
            >
              <div class="cs-chart-ref">
                <input
                  placeholder="e.g. Sheet1.B2:B7"
                  {...invalid(sliceValues)}
                  {...text(draft.series[0]?.valuesReference ?? '', (c, v) => {
                    // A part-to-whole chart only draws its first series; create it on first edit.
                    const first = c.series[0];
                    if (first) first.valuesReference = v;
                    else c.series.push(createSeries({ valuesReference: v }));
                  })}
                />
                {pick(sliceValues)}
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
              {draft.series.map((series, i) => {
                const title = data.series[i]?.name ?? `Series ${i + 1}`;
                return (
                  <SeriesRow
                    key={rowKeys.current[i]}
                    series={series}
                    title={title}
                    family={family}
                    bubble={draft.type === 'bubble'}
                    chartType={draft.type}
                    text={(value, set) => text(value, (c, v) => set(c.series[i]!, v))}
                    pick={(kind) => pick({ kind, series: i })}
                    problem={(kind) => problemOf({ kind, series: i })}
                    sample={(kind) => sampleOf({ kind, series: i })}
                    invalid={(kind) => invalid({ kind, series: i })}
                    onTypeChange={(t) =>
                      update((c) => {
                        c.series[i]!.type = t;
                      }, true)
                    }
                    onRemove={() => {
                      rowKeys.current.splice(i, 1);
                      update(
                        (c) => {
                          c.series.splice(i, 1);
                        },
                        true,
                        `Removed ${title}`,
                      );
                    }}
                  />
                );
              })}
            </div>
            <button
              type="button"
              class="cs-btn cs-chart-add"
              title={
                family === 'xy'
                  ? 'Add another set of points to plot. Give each its own X and Y values.'
                  : 'Add another set of values to plot (a second column, line, etc.). Give each its own values range.'
              }
              onClick={() => {
                rowKeys.current.push(newRowKey());
                update((c) => {
                  c.series.push(createSeries({ valuesReference: '' }));
                }, true);
              }}
            >
              + Add series
            </button>
          </>
        )}

        {ChartTypes.groupsCategories(draft.type) && (
          <Field
            label="Combine rows with the same label"
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

        <hr />

        <details
          class="cs-chart-more"
          open={draft.xAxisTitle !== null || draft.yAxisTitle !== null || !draft.showLegend}
        >
          <summary>Labels</summary>
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
        </details>
      </aside>
    </div>
  );
}

function categoryCopy(family: ChartFamily): { label: string; tip: string } {
  switch (family) {
    case 'partToWhole':
      return {
        label: 'Slice labels',
        tip: 'The cells naming each slice: one per value, e.g. Sheet1.A2:A7. Shown in the legend.',
      };
    case 'radar':
      return {
        label: 'Spoke labels',
        tip: 'The cells naming each spoke of the radar: one per value, e.g. Sheet1.A2:A7.',
      };
    case 'cartesian':
    case 'xy':
    default:
      return {
        label: 'Category labels',
        tip:
          'The cells holding the label for each data point: one per row, e.g. Sheet1.A2:A7. ' +
          'These name the columns, bars or points along the category axis.',
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

/** The reference fields a series row shows for a chart family. */
export function seriesFieldKinds(family: ChartFamily, bubble: boolean): ChartSeriesFieldKind[] {
  return [
    'nameCell',
    ...(family === 'xy' ? (['xValues'] as const) : []),
    'values',
    ...(bubble ? (['bubbleSizes'] as const) : []),
  ];
}

/**
 * One series of a range chart. Folded, it is one line (its name and values range) so a long list
 * stays short; it opens to show its fields, and stays open while one of them has a problem. A new
 * series with no values yet starts open.
 */
function SeriesRow(props: {
  series: ChartSeries;
  /** The name the legend shows. */
  title: string;
  family: ChartFamily;
  bubble: boolean;
  chartType: ChartType;
  text: TextBinder;
  /** The "Pick on sheet" button for one of this series' reference fields. */
  pick: (kind: ChartSeriesFieldKind) => ComponentChildren;
  problem: (kind: ChartSeriesFieldKind) => { readonly id: string; readonly text: string } | null;
  sample: (kind: ChartSeriesFieldKind) => string | null;
  invalid: (kind: ChartSeriesFieldKind) => Record<string, unknown>;
  onTypeChange: (type: ChartType | null) => void;
  onRemove: () => void;
}) {
  const { series, family } = props;
  const [open, setOpen] = useState(() => series.valuesReference.trim() === '');
  const hasProblem = seriesFieldKinds(family, props.bubble).some((k) => props.problem(k) !== null);
  // A problem opens the row, and it stays open once fixed: folding it while the user types in it
  // would take the field away mid-edit.
  const held = useRef(false);
  if (hasProblem) held.current = true;
  const shown = open || held.current;
  return (
    <fieldset class="cs-chart-series">
      <legend>
        <button
          type="button"
          class="cs-chart-series-toggle"
          aria-expanded={shown}
          title={shown ? 'Hide this series’ fields.' : 'Show this series’ fields.'}
          onClick={() => {
            held.current = false;
            setOpen(!shown);
          }}
        >
          <Icon name={shown ? 'chevron-down' : 'chevron-right'} />
          <span class="cs-chart-series-name">{props.title}</span>
          {!shown && <span class="cs-chart-series-ref">{series.valuesReference}</span>}
        </button>
        <button
          type="button"
          class="cs-btn danger"
          title="Remove this series from the chart. Undo brings it back."
          aria-label={`Remove ${props.title}`}
          onClick={props.onRemove}
        >
          ✕
        </button>
      </legend>
      {shown && (
        <>
          <Field
            label="Name"
            tip={'A label for this series. Leave blank to take it from a cell, or to use "Series N".'}
          >
            <input
              placeholder="Series name (optional)"
              {...props.text(series.name ?? '', (s, v) => {
                s.name = blank(v);
              })}
            />
          </Field>
          <Field
            label="Take name from cell"
            tip="A single cell whose text names the series (e.g. the column header Sheet1.B1). It wins over the typed name."
            problem={props.problem('nameCell')}
            sample={props.sample('nameCell')}
          >
            <div class="cs-chart-ref">
              <input
                placeholder="e.g. Sheet1.B1"
                {...props.invalid('nameCell')}
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
              problem={props.problem('xValues')}
              sample={props.sample('xValues')}
            >
              <div class="cs-chart-ref">
                <input
                  placeholder="e.g. Sheet1.A2:A7"
                  {...props.invalid('xValues')}
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
                ? 'The vertical position of each point: one per X value, e.g. Sheet1.B2:B7.'
                : "The cells holding this series' numbers: one per category, e.g. Sheet1.B2:B7. This is the data that gets plotted."
            }
            problem={props.problem('values')}
            sample={props.sample('values')}
          >
            <div class="cs-chart-ref">
              <input
                placeholder="e.g. Sheet1.B2:B7"
                {...props.invalid('values')}
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
              tip="Override the chart type for just this series, e.g. draw one series as a Line over Column bars to make a combo chart."
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
            <Field
              label="Bubble size"
              tip="The value that sizes each bubble: one per point, e.g. Sheet1.C2:C7."
              problem={props.problem('bubbleSizes')}
              sample={props.sample('bubbleSizes')}
            >
              <div class="cs-chart-ref">
                <input
                  placeholder="e.g. Sheet1.C2:C7"
                  {...props.invalid('bubbleSizes')}
                  {...props.text(series.bubbleSizeReference ?? '', (s, v) => {
                    s.bubbleSizeReference = blank(v);
                  })}
                />
                {props.pick('bubbleSizes')}
              </div>
            </Field>
          )}
        </>
      )}
    </fieldset>
  );
}
