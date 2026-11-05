import { useEffect, useRef, useState } from 'preact/hooks';
import type { ComponentChildren } from 'preact';
import {
  CellAddress,
  CellRange,
  FunctionRegistry,
  InputParser,
  dataValidation,
  type ColumnFilter,
  type DataValidation,
  type FilterSpec,
  type FunctionMetadata,
  type SortLevel,
  type ValidationKind,
  type ValidationOperator,
  type ValidationResponse,
} from '@cellestra/core';
import { Dialog, QuickPick } from '@cellestra/ui';
import type { AppController, DataTarget } from './controller.ts';

type Anchor = { x: number; y: number; w: number; h: number };

const close = (app: AppController) => {
  app.dialog.value = null;
  queueMicrotask(() => app.grid.current?.focus());
};

/** A small panel under an anchor rect; Escape or a click outside closes it. */
function Popover(props: {
  anchor: Anchor;
  onClose: () => void;
  children: ComponentChildren;
  class?: string;
}) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const down = (e: PointerEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) props.onClose();
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        props.onClose();
      }
    };
    addEventListener('pointerdown', down, true);
    addEventListener('keydown', key, true);
    ref.current?.querySelector<HTMLElement>('input, button, [tabindex]')?.focus();
    return () => {
      removeEventListener('pointerdown', down, true);
      removeEventListener('keydown', key, true);
    };
  }, []);
  // Keep it on screen: flip left of the anchor's right edge when it would overflow.
  const left = Math.max(4, Math.min(props.anchor.x, innerWidth - 260));
  const top = props.anchor.y + props.anchor.h + 2;
  return (
    <div ref={ref} class={`cs-popover ${props.class ?? ''}`} style={{ left: `${left}px`, top: `${top}px` }}>
      {props.children}
    </div>
  );
}

// ── Filter dropdown ─────────────────────────────────────────────────────────

export function FilterPopup(props: {
  app: AppController;
  target: DataTarget;
  column: number;
  anchor: Anchor;
}) {
  const { app, target, column } = props;
  const values = app.filterValues(target, column);
  const spec = app.sheet && currentFilter(app, target);
  const allowed = spec?.columns.get(column)?.allowedValues ?? null;
  const [checked, setChecked] = useState<Set<string>>(
    () => new Set(values.map((v) => v.key).filter((k) => !allowed || allowed.has(k))),
  );
  const [query, setQuery] = useState('');
  const header = target.hasHeader ? app.session.editText(target.range.top, column) : '';
  const name = header || `Column ${new CellAddress(column, 0).columnLabel}`;
  const shown = values.filter((v) => v.label.toLowerCase().includes(query.toLowerCase()));
  const done = () => close(app);
  const apply = (sort?: boolean) => {
    const all = values.every((v) => checked.has(v.key));
    app.applyColumnFilter(target, column, all ? null : new Set(checked), sort);
    done();
  };
  return (
    <Popover anchor={props.anchor} onClose={done} class="cs-filter-popup">
      <div class="cs-popover-title">{name}</div>
      <div class="cs-popover-row">
        <button type="button" class="cs-btn" onClick={() => apply(true)}>
          Sort A→Z
        </button>
        <button type="button" class="cs-btn" onClick={() => apply(false)}>
          Sort Z→A
        </button>
      </div>
      <input
        type="text"
        placeholder="Search values"
        value={query}
        onInput={(e) => setQuery(e.currentTarget.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') apply();
        }}
      />
      <label class="cs-check">
        <input
          type="checkbox"
          checked={shown.every((v) => checked.has(v.key))}
          onChange={(e) => {
            const next = new Set(checked);
            for (const v of shown) {
              if (e.currentTarget.checked) next.add(v.key);
              else next.delete(v.key);
            }
            setChecked(next);
          }}
        />
        (Select all)
      </label>
      <div class="cs-filter-list">
        {shown.map((v) => (
          <label key={v.key} class="cs-check">
            <input
              type="checkbox"
              checked={checked.has(v.key)}
              onChange={(e) => {
                const next = new Set(checked);
                if (e.currentTarget.checked) next.add(v.key);
                else next.delete(v.key);
                setChecked(next);
              }}
            />
            {v.label}
          </label>
        ))}
      </div>
      <div class="cs-popover-row end">
        <button
          type="button"
          class="cs-btn"
          onClick={() => {
            app.applyColumnFilter(target, column, null);
            done();
          }}
        >
          Clear
        </button>
        <button type="button" class="cs-btn primary" onClick={() => apply()}>
          Apply
        </button>
      </div>
    </Popover>
  );
}

function currentFilter(app: AppController, target: DataTarget): FilterSpec | null {
  if (target.table) return target.table.filter ?? null;
  return app.sheet.activeFilter?.equals(target.range) ? app.sheet.activeFilterSpec : null;
}

// ── Validation list dropdown ────────────────────────────────────────────────

export function ValidationListPopup(props: {
  app: AppController;
  rule: DataValidation;
  row: number;
  column: number;
  anchor: Anchor;
}) {
  const { app } = props;
  const items = app.listItems(props.rule);
  const current = app.session.editText(props.row, props.column).toUpperCase();
  const pick = (text: string) => {
    close(app);
    app.commit(props.row, props.column, text);
  };
  return (
    <Popover anchor={props.anchor} onClose={() => close(app)} class="cs-list-popup">
      {items.length === 0 && <div class="cs-muted">The list is empty.</div>}
      <div role="listbox" class="cs-filter-list">
        {items.map((t) => (
          <button
            type="button"
            role="option"
            aria-selected={t.toUpperCase() === current}
            class="cs-list-item"
            onClick={() => pick(t)}
          >
            {t}
          </button>
        ))}
      </div>
    </Popover>
  );
}

// ── Sort & Filter / Custom Sort ─────────────────────────────────────────────

export function SortFilterDialog(props: { app: AppController; target: DataTarget; sortOnly: boolean }) {
  const { app, target } = props;
  const range = target.range;
  const columns: number[] = [];
  for (let c = range.left; c <= range.right; c++) columns.push(c);
  const [hasHeader, setHasHeader] = useState(target.hasHeader);
  const label = (c: number) => {
    const h = hasHeader ? app.session.editText(range.top, c) : '';
    return h ? `${h} (${new CellAddress(c, 0).columnLabel})` : `Column ${new CellAddress(c, 0).columnLabel}`;
  };
  const existingSort = target.table
    ? target.table.sort
    : app.sheet.activeFilter?.equals(range)
      ? app.sheet.activeSort
      : null;
  const existingFilter = currentFilter(app, target);
  const [levels, setLevels] = useState<SortLevel[]>(
    existingSort?.levels.length
      ? [...existingSort.levels]
      : [{ column: app.selection.active.peek().column, ascending: true }],
  );
  const [conditions, setConditions] = useState<Map<number, string>>(
    () => new Map([...(existingFilter?.columns ?? [])].map(([c, f]) => [c, f.condition ?? ''])),
  );
  const done = () => close(app);
  const submit = () => {
    const sort = levels.length > 0 ? { levels, hasHeader } : null;
    let filter: FilterSpec | null = existingFilter;
    if (!props.sortOnly) {
      const cols = new Map<number, ColumnFilter>();
      for (const c of columns) {
        const cond = (conditions.get(c) ?? '').trim();
        const allowed = existingFilter?.columns.get(c)?.allowedValues ?? null;
        if (cond || allowed) cols.set(c, { allowedValues: allowed, condition: cond || null });
      }
      filter = cols.size > 0 ? { columns: cols, hasHeader } : null;
    }
    app.sortFilterCommit(target, sort, filter);
    done();
  };
  const setLevel = (i: number, l: Partial<SortLevel>) =>
    setLevels(levels.map((x, j) => (j === i ? { ...x, ...l } : x)));
  return (
    <Dialog
      title={props.sortOnly ? 'Custom Sort' : 'Sort & Filter'}
      class="cs-data-dialog"
      onClose={done}
      onSubmit={submit}
      actions={
        <>
          <button type="button" class="cs-btn" onClick={done}>
            Cancel
          </button>
          <button type="submit" class="cs-btn primary">
            OK
          </button>
        </>
      }
    >
      <p class="cs-muted">
        {target.table ? `Table ${target.table.name}` : 'Range'} {range.toString()}
      </p>
      <label class="cs-check">
        <input type="checkbox" checked={hasHeader} onChange={(e) => setHasHeader(e.currentTarget.checked)} />
        First row is a header
      </label>
      <h3>Sort by</h3>
      {levels.map((l, i) => (
        <div class="cs-form-row" key={i}>
          <span class="cs-muted">{i === 0 ? 'Sort by' : 'Then by'}</span>
          <select value={l.column} onChange={(e) => setLevel(i, { column: Number(e.currentTarget.value) })}>
            {columns.map((c) => (
              <option value={c}>{label(c)}</option>
            ))}
          </select>
          <select
            value={l.ascending ? 'asc' : 'desc'}
            onChange={(e) => setLevel(i, { ascending: e.currentTarget.value === 'asc' })}
          >
            <option value="asc">Ascending</option>
            <option value="desc">Descending</option>
          </select>
          <button
            type="button"
            class="cs-btn"
            aria-label="Remove level"
            onClick={() => setLevels(levels.filter((_, j) => j !== i))}
          >
            ×
          </button>
        </div>
      ))}
      {levels.length < 3 && (
        <button
          type="button"
          class="cs-btn"
          onClick={() =>
            setLevels([
              ...levels,
              {
                column: columns.find((c) => !levels.some((l) => l.column === c)) ?? range.left,
                ascending: true,
              },
            ])
          }
        >
          Add level
        </button>
      )}
      {!props.sortOnly && (
        <>
          <h3>Filter</h3>
          <p class="cs-muted">Conditions work like SUMIF criteria: &gt;5, &lt;&gt;0, *text*, =Done.</p>
          <div class="cs-filter-grid">
            {columns.map((c) => (
              <label key={c}>
                {label(c)}
                <input
                  type="text"
                  value={conditions.get(c) ?? ''}
                  onInput={(e) => setConditions(new Map(conditions).set(c, e.currentTarget.value))}
                />
              </label>
            ))}
          </div>
        </>
      )}
    </Dialog>
  );
}

// ── Data validation ─────────────────────────────────────────────────────────

const KINDS: [ValidationKind, string][] = [
  ['any', 'Any value'],
  ['list', 'List'],
  ['wholeNumber', 'Whole number'],
  ['decimal', 'Decimal'],
  ['date', 'Date'],
  ['textLength', 'Text length'],
];
const OPERATORS: [ValidationOperator, string][] = [
  ['between', 'between'],
  ['notBetween', 'not between'],
  ['equal', 'equal to'],
  ['notEqual', 'not equal to'],
  ['greaterThan', 'greater than'],
  ['lessThan', 'less than'],
  ['greaterOrEqual', 'greater than or equal to'],
  ['lessOrEqual', 'less than or equal to'],
];
const RESPONSES: [ValidationResponse, string][] = [
  ['stop', 'Stop — reject the entry'],
  ['warning', 'Warning — allow, but say so'],
  ['information', 'Information — allow, but say so'],
];

export function ValidationDialog(props: {
  app: AppController;
  range: CellRange;
  existing: DataValidation | null;
}) {
  const { app } = props;
  const e = props.existing;
  const [kind, setKind] = useState<ValidationKind>(e?.kind ?? 'list');
  const [op, setOp] = useState<ValidationOperator>(e?.operator ?? 'between');
  const [a, setA] = useState(e && e.kind !== 'list' && e.kind !== 'any' ? String(e.operand1) : '');
  const [b, setB] = useState(e && e.kind !== 'list' && e.kind !== 'any' ? String(e.operand2) : '');
  const [list, setList] = useState(e?.listValues?.join(', ') ?? '');
  const [source, setSource] = useState(e?.listSource?.toString() ?? '');
  const [input, setInput] = useState(e?.inputMessage ?? '');
  const [error, setError] = useState(e?.errorMessage ?? '');
  const [response, setResponse] = useState<ValidationResponse>(e?.response ?? 'stop');
  const [problem, setProblem] = useState('');
  const target = e?.target ?? props.range;
  const between = op === 'between' || op === 'notBetween';
  const numeric = kind !== 'any' && kind !== 'list';
  const done = () => close(app);
  const operand = (text: string) => {
    const n = Number(text);
    if (text.trim() !== '' && Number.isFinite(n)) return n;
    if (kind === 'date' && text.trim()) {
      const v = InputParser.parse(text, app.session.locale).value;
      if (v.kind === 'number') return v.value;
    }
    return null;
  };
  const submit = () => {
    let rest: Partial<DataValidation> = {
      operator: op,
      response,
      inputMessage: input.trim() || null,
      errorMessage: error.trim() || null,
    };
    if (kind === 'list') {
      const src = source.trim() ? CellRange.tryParse(source.trim().replace(/\$/g, '')) : null;
      if (source.trim() && !src) return setProblem('The source isn’t a range like A1:A10.');
      const values = list
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
      if (!src && values.length === 0) return setProblem('Type some list items or a source range.');
      rest = src ? { ...rest, listSource: src, listValues: null } : { ...rest, listValues: values };
    } else if (numeric) {
      const x = operand(a);
      const y = between ? operand(b) : 0;
      if (x === null || y === null) return setProblem('Enter a number (or a date) for each bound.');
      rest = { ...rest, operand1: x, operand2: y };
    }
    app.setValidation(target, dataValidation(kind, target, rest));
    done();
  };
  return (
    <Dialog
      title="Data Validation"
      class="cs-data-dialog"
      onClose={done}
      onSubmit={submit}
      actions={
        <>
          {e && (
            <button
              type="button"
              class="cs-btn danger"
              onClick={() => {
                app.setValidation(target, null);
                done();
              }}
            >
              Remove Rule
            </button>
          )}
          <button type="button" class="cs-btn" onClick={done}>
            Cancel
          </button>
          <button type="submit" class="cs-btn primary">
            OK
          </button>
        </>
      }
    >
      <p class="cs-muted">Applies to {target.toString()}</p>
      <label>
        Allow
        <select value={kind} onChange={(ev) => setKind(ev.currentTarget.value as ValidationKind)}>
          {KINDS.map(([k, t]) => (
            <option value={k}>{t}</option>
          ))}
        </select>
      </label>
      {kind === 'list' && (
        <>
          <label>
            Items (comma-separated)
            <input type="text" value={list} onInput={(ev) => setList(ev.currentTarget.value)} />
          </label>
          <label>
            …or take them from a range
            <input
              type="text"
              placeholder="e.g. H2:H20"
              value={source}
              onInput={(ev) => setSource(ev.currentTarget.value)}
            />
          </label>
        </>
      )}
      {numeric && (
        <div class="cs-form-row">
          <select value={op} onChange={(ev) => setOp(ev.currentTarget.value as ValidationOperator)}>
            {OPERATORS.map(([k, t]) => (
              <option value={k}>{t}</option>
            ))}
          </select>
          <input type="text" aria-label="Value" value={a} onInput={(ev) => setA(ev.currentTarget.value)} />
          {between && (
            <>
              <span class="cs-muted">and</span>
              <input
                type="text"
                aria-label="Second value"
                value={b}
                onInput={(ev) => setB(ev.currentTarget.value)}
              />
            </>
          )}
        </div>
      )}
      <label>
        Input message (shown when the cell is selected)
        <input type="text" value={input} onInput={(ev) => setInput(ev.currentTarget.value)} />
      </label>
      <label>
        Error message
        <input type="text" value={error} onInput={(ev) => setError(ev.currentTarget.value)} />
      </label>
      <label>
        When the entry is invalid
        <select value={response} onChange={(ev) => setResponse(ev.currentTarget.value as ValidationResponse)}>
          {RESPONSES.map(([k, t]) => (
            <option value={k}>{t}</option>
          ))}
        </select>
      </label>
      {problem && <p class="cs-problem">{problem}</p>}
    </Dialog>
  );
}

// ── Hyperlink ───────────────────────────────────────────────────────────────

export function HyperlinkDialog(props: { app: AppController; row: number; column: number }) {
  const { app, row, column } = props;
  const existing = app.sheet.find(row, column)?.hyperlink ?? null;
  const [url, setUrl] = useState(existing?.target ?? '');
  const [text, setText] = useState(app.session.editText(row, column));
  const [tooltip, setTooltip] = useState(existing?.tooltip ?? '');
  const done = () => close(app);
  const submit = () => {
    const target = url.trim();
    if (!target) {
      if (existing) app.setHyperlink(row, column, null, text);
      return done();
    }
    const normalized = /^[a-z][a-z0-9+.-]*:|^#/i.test(target) ? target : `https://${target}`;
    app.setHyperlink(
      row,
      column,
      { target: normalized, tooltip: tooltip.trim() || null },
      text || normalized,
    );
    done();
  };
  return (
    <Dialog
      title="Hyperlink"
      class="cs-data-dialog"
      onClose={done}
      onSubmit={submit}
      actions={
        <>
          {existing && (
            <button
              type="button"
              class="cs-btn danger"
              onClick={() => {
                app.setHyperlink(row, column, null, text);
                done();
              }}
            >
              Remove Link
            </button>
          )}
          <button type="button" class="cs-btn" onClick={done}>
            Cancel
          </button>
          <button type="submit" class="cs-btn primary">
            OK
          </button>
        </>
      }
    >
      <label>
        Address
        <input
          type="text"
          placeholder="https://example.com  or  #Sheet2.A1"
          value={url}
          autofocus
          onInput={(e) => setUrl(e.currentTarget.value)}
        />
      </label>
      <label>
        Text to show
        <input type="text" value={text} onInput={(e) => setText(e.currentTarget.value)} />
      </label>
      <label>
        Tooltip
        <input type="text" value={tooltip} onInput={(e) => setTooltip(e.currentTarget.value)} />
      </label>
      <p class="cs-muted">⌘-click (Ctrl-click) a linked cell to open it.</p>
    </Dialog>
  );
}

// ── Named ranges ────────────────────────────────────────────────────────────

export function NamedRangesDialog({ app }: { app: AppController }) {
  const wb = app.session.workbook.value;
  void app.session.version.value;
  const r = app.range;
  const [name, setName] = useState('');
  const [ref, setRef] = useState(r.toString());
  const [scope, setScope] = useState('');
  const [problem, setProblem] = useState('');
  const done = () => close(app);
  const add = () => {
    const target =
      CellRange.tryParse(ref.trim().replace(/\$/g, '')) ??
      (CellAddress.tryParse(ref.trim().replace(/\$/g, ''))
        ? new CellRange(CellAddress.parse(ref.trim().replace(/\$/g, '')))
        : null);
    if (!name.trim()) return setProblem('Type a name.');
    if (!target) return setProblem('The range isn’t like A1 or B2:D9.');
    const err = app.addNamedRange({ name: name.trim(), target, scope: scope || null });
    if (err) return setProblem(err);
    setProblem('');
    setName('');
  };
  return (
    <Dialog
      title="Named Ranges"
      class="cs-data-dialog"
      onClose={done}
      onSubmit={add}
      actions={
        <button type="button" class="cs-btn primary" onClick={done}>
          Done
        </button>
      }
    >
      {wb.namedRanges.length === 0 ? (
        <p class="cs-muted">No names yet. Formulas can use a name in place of a range, e.g. =SUM(Sales).</p>
      ) : (
        <table class="cs-names">
          <tbody>
            {wb.namedRanges.map((n) => (
              <tr key={`${n.scope ?? ''}!${n.name}`}>
                <td>{n.name}</td>
                <td>{n.target.toString()}</td>
                <td class="cs-muted">{n.scope ?? 'Workbook'}</td>
                <td>
                  <button
                    type="button"
                    class="cs-btn"
                    aria-label={`Delete ${n.name}`}
                    onClick={() => app.removeNamedRange(n)}
                  >
                    Delete
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
      <h3>New name</h3>
      <div class="cs-form-row">
        <input
          type="text"
          aria-label="Name"
          placeholder="Name"
          value={name}
          autofocus
          onInput={(e) => setName(e.currentTarget.value)}
        />
        <input type="text" aria-label="Range" value={ref} onInput={(e) => setRef(e.currentTarget.value)} />
        <select aria-label="Scope" value={scope} onChange={(e) => setScope(e.currentTarget.value)}>
          <option value="">Workbook</option>
          {app.session.sheetNames.value.map((s) => (
            <option value={s}>{s}</option>
          ))}
        </select>
        <button type="submit" class="cs-btn">
          Add
        </button>
      </div>
      {problem && <p class="cs-problem">{problem}</p>}
    </Dialog>
  );
}

// ── Insert function ─────────────────────────────────────────────────────────

let functions: FunctionMetadata[] | null = null;

export function FunctionPicker({ app }: { app: AppController }) {
  functions ??= FunctionRegistry.createDefault().allMetadata();
  return (
    <QuickPick<FunctionMetadata>
      items={functions}
      label={(f) => `${f.signature} — ${f.summary}`}
      placeholder="Insert function…"
      onPick={(f) => queueMicrotask(() => app.grid.current?.beginEdit(`=${f.name}(`))}
      onClose={() => {
        app.dialog.value = null;
        queueMicrotask(() => app.grid.current?.focus());
      }}
    />
  );
}
