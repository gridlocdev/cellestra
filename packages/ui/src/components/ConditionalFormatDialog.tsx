import {
  Argb,
  FORMAT_PRESETS,
  RULE_KINDS,
  applyPreset,
  editRule,
  newRule,
  rulesOf,
  rulesTransform,
  type CellStyle,
  type FormatPresetId,
  type RuleDraft,
  type RuleKind,
} from '@cellestra/core';
import { useMemo, useState } from 'preact/hooks';
import { Dialog } from './Dialog.tsx';

/**
 * Conditional formatting for the selection: rules checked in order, the first that holds giving the
 * cell its format. Seeded from the active cell; OK gives every selected cell these rules, with
 * relative references in them relative to `base` (the selection's top-left cell).
 */
export function ConditionalFormatDialog(props: {
  readonly current: CellStyle;
  readonly base: { readonly sheet: string; readonly row: number; readonly column: number };
  readonly onApply: (transform: (s: CellStyle) => CellStyle) => void;
  readonly onClose: () => void;
}) {
  const [rules, setRules] = useState<RuleDraft[]>(() => rulesOf(props.current));
  const [tried, setTried] = useState(false);
  const result = useMemo(() => rulesTransform(rules, props.base), [rules, props.base]);
  const error = 'error' in result ? result.error : null;
  const update = (i: number, patch: Parameters<typeof editRule>[1]) =>
    setRules((rs) => rs.map((r, j) => (j === i ? editRule(r, patch) : r)));
  const move = (i: number, by: number) =>
    setRules((rs) => {
      const next = [...rs];
      const [r] = next.splice(i, 1);
      next.splice(i + by, 0, r!);
      return next;
    });

  const submit = () => {
    setTried(true);
    if ('error' in result) return;
    props.onApply(result.transform);
    props.onClose();
  };

  return (
    <Dialog
      title="Conditional Formatting"
      class="cs-conditional"
      onClose={props.onClose}
      onSubmit={submit}
      actions={
        <>
          <button type="button" class="cs-btn" onClick={props.onClose}>
            Cancel
          </button>
          <button type="submit" class="cs-btn primary">
            OK
          </button>
        </>
      }
    >
      <p class="cs-cf-note">
        Rules for the selected cells, checked in order: the first one that holds formats the cell. References
        like B1 are relative to the top-left selected cell.
      </p>
      {rules.length === 0 && <p class="cs-cf-empty">No rules. Add one to start.</p>}
      <ol class="cs-cf-rules">
        {rules.map((rule, i) => (
          <RuleRow
            key={i}
            index={i}
            rule={rule}
            plain={props.current}
            first={i === 0}
            last={i === rules.length - 1}
            onChange={(patch) => update(i, patch)}
            onMove={(by) => move(i, by)}
            onRemove={() => setRules((rs) => rs.filter((_, j) => j !== i))}
          />
        ))}
      </ol>
      <button type="button" class="cs-btn" onClick={() => setRules((rs) => [...rs, newRule()])}>
        Add rule
      </button>
      {tried && error && (
        <p class="cs-cf-error" role="alert">
          {error}
        </p>
      )}
    </Dialog>
  );
}

function RuleRow(props: {
  readonly index: number;
  readonly rule: RuleDraft;
  readonly plain: CellStyle;
  readonly first: boolean;
  readonly last: boolean;
  readonly onChange: (patch: Parameters<typeof editRule>[1]) => void;
  readonly onMove: (by: number) => void;
  readonly onRemove: () => void;
}) {
  const { rule, index: i } = props;
  const operands = RULE_KINDS.find((k) => k.kind === rule.kind)!.operands;
  const sample =
    rule.format === 'keep' ? (rule.keepStyle ?? props.plain) : applyPreset(rule.format, props.plain);
  return (
    <li class="cs-cf-rule">
      <div class="cs-cf-line">
        <select
          aria-label={`Rule ${i + 1} condition`}
          value={rule.kind}
          onChange={(e) => props.onChange({ kind: (e.target as HTMLSelectElement).value as RuleKind })}
        >
          {RULE_KINDS.map((k) => (
            <option key={k.kind} value={k.kind}>
              {k.label}
            </option>
          ))}
        </select>
        <input
          aria-label={`Rule ${i + 1} ${rule.kind === 'formula' ? 'formula' : 'value'}`}
          placeholder={rule.kind === 'formula' ? 'A1>B1' : '10, B1 or text'}
          value={rule.value1}
          onInput={(e) => props.onChange({ value1: (e.target as HTMLInputElement).value })}
        />
        {operands === 2 && (
          <>
            <span>and</span>
            <input
              aria-label={`Rule ${i + 1} second value`}
              value={rule.value2}
              onInput={(e) => props.onChange({ value2: (e.target as HTMLInputElement).value })}
            />
          </>
        )}
      </div>
      <div class="cs-cf-line">
        <select
          aria-label={`Rule ${i + 1} format`}
          value={rule.format}
          onChange={(e) =>
            props.onChange({ format: (e.target as HTMLSelectElement).value as FormatPresetId })
          }
        >
          {FORMAT_PRESETS.map((p) => (
            <option key={p.id} value={p.id}>
              {p.label}
            </option>
          ))}
          {rule.keepStyle !== undefined && <option value="keep">The file's own format</option>}
        </select>
        <span
          class="cs-cf-sample"
          style={{
            background: sample.fill.background !== null ? cssColor(sample.fill.background) : 'transparent',
            color: cssColor(sample.font.color),
            fontWeight: sample.font.bold ? 'bold' : 'normal',
          }}
        >
          AaBb 123
        </span>
        <span class="cs-cf-tools">
          <button
            type="button"
            class="cs-btn"
            disabled={props.first}
            aria-label="Move up"
            onClick={() => props.onMove(-1)}
          >
            ↑
          </button>
          <button
            type="button"
            class="cs-btn"
            disabled={props.last}
            aria-label="Move down"
            onClick={() => props.onMove(1)}
          >
            ↓
          </button>
          <button type="button" class="cs-btn" onClick={props.onRemove}>
            Remove
          </button>
        </span>
      </div>
    </li>
  );
}

const cssColor = (argb: number) => (Argb.isTransparent(argb) ? 'inherit' : Argb.toCss(argb));
