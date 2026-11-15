import type { ChartDefinition } from '@cellestra/charts';
import { ChartTableBinding, type TableColumnInfo, type Workbook } from '@cellestra/core';
import type { ComponentChildren } from 'preact';
import { useMemo } from 'preact/hooks';
import type { ChartFamily } from './ChartSheetView.tsx';

export interface ChartTableFieldsProps {
  readonly workbook: Workbook;
  /** Bumped whenever the workbook changes, so the column list follows the table. */
  readonly version: number;
  readonly chart: ChartDefinition;
  readonly family: ChartFamily;
  /** A committed edit. */
  readonly onChange: (next: ChartDefinition) => void;
}

function Picker(props: { label: string; tip: string; children: ComponentChildren }) {
  return (
    <label class="cs-chart-field" title={props.tip}>
      <span>{props.label}</span>
      {props.children}
    </label>
  );
}

function ColumnSelect(props: {
  columns: readonly TableColumnInfo[];
  value: string | null;
  none: string | null;
  onChange: (column: string | null) => void;
}) {
  return (
    <select
      value={props.value ?? ''}
      onChange={(e) => {
        const v = (e.currentTarget as HTMLSelectElement).value;
        props.onChange(v === '' ? null : v);
      }}
    >
      {(props.none !== null || props.value === null) && (
        <option value="">{props.none ?? 'Choose a column…'}</option>
      )}
      {props.columns.map((c) => (
        <option key={c.column} value={c.name}>
          {c.name}
        </option>
      ))}
    </select>
  );
}

const CATEGORY_COPY: Record<ChartFamily, { label: string; tip: string }> = {
  cartesian: {
    label: 'X / Categories',
    tip: 'The column that labels each point along the category axis. Rows with the same label become one category.',
  },
  radar: {
    label: 'Spoke labels',
    tip: 'The column that names each spoke. Rows with the same label become one spoke.',
  },
  partToWhole: {
    label: 'Slice labels',
    tip: 'The column that names each slice. Rows with the same label become one slice.',
  },
  xy: { label: 'X values', tip: 'The column holding the horizontal position of each point.' },
};

const same = (a: string, b: string): boolean => a.toLowerCase() === b.toLowerCase();

/**
 * The chart panel's table mode: plot a named table by picking its columns. Every pick is one
 * committed edit; the rules (what a type can plot, keeping series settings) live in
 * `ChartTableBinding`.
 */
export function ChartTableFields(props: ChartTableFieldsProps) {
  const { workbook, chart, family } = props;
  const source = chart.source;
  const tables = useMemo(() => ChartTableBinding.tableNames(workbook), [workbook, props.version]);
  const columns = useMemo(
    () => (source ? ChartTableBinding.columns(workbook, source.table) : null),
    [workbook, source?.table, props.version],
  );
  const plotted = source?.series.map((s) => s.column) ?? [];
  const isPlotted = (name: string) => plotted.some((p) => same(p, name));

  const tablePicker = (
    <Picker
      label="Table"
      tip="The named table the chart plots. The chart follows it as rows are added or removed."
    >
      <select
        value={source?.table ?? ''}
        onChange={(e) => {
          const name = (e.currentTarget as HTMLSelectElement).value;
          if (!name) return;
          if (source) props.onChange(ChartTableBinding.setTable(workbook, chart, name));
          else {
            const result = ChartTableBinding.toTable(workbook, chart, name);
            if (result.ok) props.onChange(result.chart);
          }
        }}
      >
        {!source && <option value="">Choose a table…</option>}
        {source && !tables.some((t) => same(t, source.table)) && (
          <option value={source.table}>{source.table}</option>
        )}
        {tables.map((t) => (
          <option key={t} value={t}>
            {t}
          </option>
        ))}
      </select>
    </Picker>
  );

  if (!source) return tablePicker;
  if (!columns)
    return (
      <>
        {tablePicker}
        <p class="cs-chart-hint" role="status">
          There is no table called “{source.table}” any more, so the chart shows the cells it last plotted.
          Pick another table, or switch to Range.
        </p>
      </>
    );

  const category = CATEGORY_COPY[family];
  const values = columns.filter(
    (c) => source.categoryColumn === null || !same(c.name, source.categoryColumn),
  );
  const setSeries = (names: readonly string[]) =>
    props.onChange(ChartTableBinding.setSeriesColumns(workbook, chart, names));

  return (
    <>
      {tablePicker}
      <Picker label={category.label} tip={category.tip}>
        <ColumnSelect
          columns={columns}
          value={source.categoryColumn}
          none={family === 'xy' ? null : 'None (number the rows)'}
          onChange={(column) => props.onChange(ChartTableBinding.setCategoryColumn(workbook, chart, column))}
        />
      </Picker>

      {family === 'partToWhole' || chart.type === 'bubble' ? (
        <Picker
          label={family === 'partToWhole' ? 'Slice values' : 'Y values'}
          tip={
            family === 'partToWhole'
              ? 'The column whose numbers size the slices. A pie or doughnut shows one column.'
              : 'The column holding the vertical position of each bubble.'
          }
        >
          <ColumnSelect
            columns={values}
            value={plotted[0] ?? null}
            none={null}
            onChange={(column) => setSeries(column === null ? [] : [column])}
          />
        </Picker>
      ) : (
        <fieldset class="cs-chart-columns" title="Tick each column to plot as its own series.">
          <legend>{family === 'xy' ? 'Y values' : 'Values'}</legend>
          {values.length === 0 && <p class="cs-chart-hint">The table has no other columns to plot.</p>}
          {values.map((c) => (
            <label key={c.column} class="cs-chart-check">
              <input
                type="checkbox"
                checked={isPlotted(c.name)}
                onChange={(e) => {
                  const on = (e.currentTarget as HTMLInputElement).checked;
                  setSeries(on ? [...plotted, c.name] : plotted.filter((p) => !same(p, c.name)));
                }}
              />
              {c.name}
              {c.kind !== 'number' && (
                <span class="cs-chart-kind">{c.kind === 'empty' ? 'empty' : 'text'}</span>
              )}
            </label>
          ))}
        </fieldset>
      )}

      {chart.type === 'bubble' && (
        <Picker label="Bubble size" tip="The column whose numbers size each bubble.">
          <ColumnSelect
            columns={columns}
            value={source.series[0]?.sizeColumn ?? null}
            none={null}
            onChange={(column) => props.onChange(ChartTableBinding.setSizeColumn(workbook, chart, column))}
          />
        </Picker>
      )}
    </>
  );
}
