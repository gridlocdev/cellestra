import {
  Argb,
  type ArgbColor,
  type CellStyle,
  type HorizontalAlignment,
  type VerticalAlignment,
} from '@cellestra/core';
import type { CSSProperties } from 'preact';
import { useEffect, useMemo, useRef, useState } from 'preact/hooks';
import { Dialog } from './Dialog.tsx';
import {
  BORDER_LINE_STYLES,
  BORDER_PRESETS,
  FONT_SIZE_MAX,
  FONT_SIZE_MIN,
  HORIZONTAL_ALIGNMENTS,
  NUMBER_FORMAT_PRESETS,
  VARIES,
  VERTICAL_ALIGNMENTS,
  buildFormatTransform,
  editFormatCells,
  formatSample,
  nextTriState,
  parseFontSize,
  presetIndex,
  seedFormatCells,
  type BorderEdgesOn,
  type FormatCellsValues,
} from './formatCells.ts';

const CUSTOM = 'custom';
const VARIES_VALUE = 'varies';

/**
 * The Format Cells dialog (PRD §6.2, ⌘1): number format, font emphasis/size/colour, fill, alignment,
 * and borders. `styles` are the distinct styles across the selection; fields they disagree on show
 * "(varies)" and stay per-cell unless touched. `onApply` gets a transform for every selected cell.
 */
export function FormatCellsDialog(props: {
  readonly current: CellStyle;
  readonly styles: readonly CellStyle[];
  readonly locale: string;
  readonly onApply: (transform: (s: CellStyle) => CellStyle) => void;
  readonly onClose: () => void;
}) {
  const [state, setState] = useState(() => seedFormatCells(props.current, props.styles));
  const { values: v, varies } = state;
  const edit = (patch: Partial<FormatCellsValues>) => setState((s) => editFormatCells(s, patch));
  const transform = useMemo(() => buildFormatTransform(state), [state]);
  const preview = useMemo(() => transform(props.current), [transform, props.current]);

  const submit = () => {
    props.onApply(transform);
    props.onClose();
  };

  return (
    <Dialog
      title="Format Cells"
      class="cs-format-cells"
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
      <NumberSection
        code={v.numberFormat}
        locale={props.locale}
        onCode={(numberFormat) => edit({ numberFormat })}
      />

      <fieldset class="cs-fc-section">
        <legend>Font</legend>
        <div class="cs-fc-checks">
          {(['bold', 'italic', 'underline', 'strikethrough'] as const).map((f) => (
            <TriCheck
              key={f}
              label={f[0]!.toUpperCase() + f.slice(1)}
              value={v[f]}
              onClick={() => edit({ [f]: nextTriState(v[f], varies[f]) })}
            />
          ))}
        </div>
        <div class="cs-fc-row">
          <label for="cs-fc-size">Size</label>
          <input
            id="cs-fc-size"
            type="number"
            min={FONT_SIZE_MIN}
            max={FONT_SIZE_MAX}
            step={0.5}
            placeholder={varies.size ? VARIES : undefined}
            value={v.size ?? ''}
            onChange={(e) => edit({ size: parseFontSize((e.target as HTMLInputElement).value) })}
          />
        </div>
        <div class="cs-fc-row">
          <label for="cs-fc-text-color">Text color</label>
          <ColorInput
            id="cs-fc-text-color"
            value={v.textColor}
            varies={varies.textColor && !state.touched.has('textColor')}
            onChange={(textColor) => edit({ textColor })}
          />
          <button
            type="button"
            class="cs-btn"
            title="Reset the text colour to black"
            onClick={() => edit({ textColor: Argb.black })}
          >
            Reset
          </button>
        </div>
      </fieldset>

      <fieldset class="cs-fc-section">
        <legend>Fill</legend>
        <div class="cs-fc-row">
          <TriCheck
            label="Fill"
            value={v.fill}
            onClick={() => edit({ fill: nextTriState(v.fill, varies.fill) })}
          />
          <ColorInput
            id="cs-fc-fill-color"
            label="Fill color"
            value={v.fillColor}
            disabled={v.fill !== true}
            varies={varies.fillColor && !state.touched.has('fillColor')}
            onChange={(fillColor) => edit({ fillColor })}
          />
          <button
            type="button"
            class="cs-btn"
            title="Clear the fill (no background)"
            onClick={() => edit({ fill: false, fillColor: Argb.white })}
          >
            Reset
          </button>
        </div>
      </fieldset>

      <fieldset class="cs-fc-section">
        <legend>Alignment</legend>
        <div class="cs-fc-row">
          <label for="cs-fc-halign">Horizontal</label>
          <VariesSelect
            id="cs-fc-halign"
            options={HORIZONTAL_ALIGNMENTS}
            value={v.horizontal}
            varies={varies.horizontal}
            onChange={(horizontal: HorizontalAlignment | null) => edit({ horizontal })}
          />
        </div>
        <div class="cs-fc-row">
          <label for="cs-fc-valign">Vertical</label>
          <VariesSelect
            id="cs-fc-valign"
            options={VERTICAL_ALIGNMENTS}
            value={v.vertical}
            varies={varies.vertical}
            onChange={(vertical: VerticalAlignment | null) => edit({ vertical })}
          />
        </div>
        <TriCheck
          label="Wrap text"
          value={v.wrapText}
          onClick={() => edit({ wrapText: nextTriState(v.wrapText, varies.wrapText) })}
        />
      </fieldset>

      <fieldset class="cs-fc-section">
        <legend>Border{varies.borderEdges && !state.touched.has('borderEdges') ? ` ${VARIES}` : ''}</legend>
        <div class="cs-fc-row">
          <label for="cs-fc-border-style">Line</label>
          <select
            id="cs-fc-border-style"
            value={v.borderStyle}
            onChange={(e) =>
              edit({ borderStyle: (e.target as HTMLSelectElement).value as FormatCellsValues['borderStyle'] })
            }
          >
            {BORDER_LINE_STYLES.map((b) => (
              <option key={b.style} value={b.style}>
                {b.label}
              </option>
            ))}
          </select>
          <ColorInput
            id="cs-fc-border-color"
            label="Border color"
            value={v.borderColor}
            onChange={(borderColor) => edit({ borderColor })}
          />
        </div>
        <div class="cs-fc-checks">
          {(['top', 'bottom', 'left', 'right'] as const).map((side) => (
            <label key={side} class="cs-fc-check">
              <input
                type="checkbox"
                checked={v.borderEdges[side]}
                onChange={(e) =>
                  edit({ borderEdges: { ...v.borderEdges, [side]: (e.target as HTMLInputElement).checked } })
                }
              />
              {side[0]!.toUpperCase() + side.slice(1)}
            </label>
          ))}
        </div>
        <div class="cs-fc-presets">
          {BORDER_PRESETS.map((p) => (
            <button
              key={p.label}
              type="button"
              class="cs-btn"
              aria-pressed={sameEdges(p.edges, v.borderEdges)}
              onClick={() => edit({ borderEdges: p.edges })}
            >
              {p.label}
            </button>
          ))}
        </div>
      </fieldset>

      <div class="cs-fc-preview" aria-label="Preview">
        <div class="cs-fc-preview-cell" style={previewCss(preview)}>
          {formatSample(preview.numberFormat.code, props.locale).text || ' '}
        </div>
      </div>
    </Dialog>
  );
}

/** Preset menu + editable code + live sample. `code` null = the selection's formats vary. */
function NumberSection(props: {
  readonly code: string | null;
  readonly locale: string;
  readonly onCode: (code: string) => void;
}) {
  const { code } = props;
  const index = code === null ? -1 : presetIndex(code);
  const selectValue = code === null ? VARIES_VALUE : index < 0 ? CUSTOM : String(index);
  const sample = code === null ? null : formatSample(code, props.locale);
  return (
    <fieldset class="cs-fc-section">
      <legend>Number</legend>
      <div class="cs-fc-row">
        <label for="cs-fc-format">Category</label>
        <select
          id="cs-fc-format"
          value={selectValue}
          onChange={(e) => {
            const preset = NUMBER_FORMAT_PRESETS[Number((e.target as HTMLSelectElement).value)];
            if (preset) props.onCode(preset.code);
          }}
        >
          {NUMBER_FORMAT_PRESETS.map((p, i) => (
            <option key={p.code} value={String(i)}>
              {p.label}
            </option>
          ))}
          {selectValue === CUSTOM && <option value={CUSTOM}>Custom</option>}
          {selectValue === VARIES_VALUE && (
            <option value={VARIES_VALUE} disabled>
              {VARIES}
            </option>
          )}
        </select>
      </div>
      <div class="cs-fc-row">
        <label for="cs-fc-code">Code</label>
        <input
          id="cs-fc-code"
          type="text"
          class="cs-fc-code"
          spellcheck={false}
          placeholder={code === null ? VARIES : undefined}
          value={code ?? ''}
          onInput={(e) => props.onCode((e.target as HTMLInputElement).value)}
        />
      </div>
      <div class="cs-fc-row">
        <span class="cs-fc-label">Sample</span>
        <output class="cs-fc-sample" for="cs-fc-code">
          {sample === null ? VARIES : sample.text}
          {sample?.negative !== undefined && <span class="cs-fc-negative">{sample.negative}</span>}
        </output>
      </div>
    </fieldset>
  );
}

/** A checkbox that shows `null` as indeterminate; clicks are cycled by the caller. */
function TriCheck(props: {
  readonly label: string;
  readonly value: boolean | null;
  readonly onClick: () => void;
}) {
  const ref = useRef<HTMLInputElement>(null);
  useEffect(() => {
    if (ref.current) ref.current.indeterminate = props.value === null;
  }, [props.value]);
  return (
    <label class="cs-fc-check">
      <input
        ref={ref}
        type="checkbox"
        checked={props.value === true}
        aria-checked={props.value === null ? 'mixed' : props.value}
        onClick={(e) => {
          e.preventDefault();
          props.onClick();
        }}
      />
      {props.label}
    </label>
  );
}

function ColorInput(props: {
  readonly id: string;
  readonly label?: string;
  readonly value: ArgbColor;
  readonly varies?: boolean;
  readonly disabled?: boolean;
  readonly onChange: (color: ArgbColor) => void;
}) {
  return (
    <span class={`cs-fc-color${props.varies ? ' varies' : ''}`} title={props.varies ? VARIES : undefined}>
      <input
        id={props.id}
        type="color"
        aria-label={props.label}
        disabled={props.disabled}
        value={Argb.toHex(props.value).toLowerCase()}
        onInput={(e) => {
          const c = Argb.tryFromHex((e.target as HTMLInputElement).value);
          if (c !== null) props.onChange(c);
        }}
      />
    </span>
  );
}

/** A select with an extra "(varies)" entry while the selection disagrees and nothing is chosen. */
function VariesSelect<T extends string>(props: {
  readonly id: string;
  readonly options: readonly { readonly label: string; readonly value: T }[];
  readonly value: T | null;
  readonly varies: boolean;
  readonly onChange: (value: T | null) => void;
}) {
  return (
    <select
      id={props.id}
      value={props.value ?? VARIES_VALUE}
      onChange={(e) => {
        const raw = (e.target as HTMLSelectElement).value;
        props.onChange(raw === VARIES_VALUE ? null : (raw as T));
      }}
    >
      {props.options.map((o) => (
        <option key={o.value} value={o.value}>
          {o.label}
        </option>
      ))}
      {props.varies && <option value={VARIES_VALUE}>{VARIES}</option>}
    </select>
  );
}

function sameEdges(a: BorderEdgesOn, b: BorderEdgesOn): boolean {
  return a.top === b.top && a.bottom === b.bottom && a.left === b.left && a.right === b.right;
}

const CSS_BORDER: Record<string, string> = {
  thin: '1px solid',
  medium: '2px solid',
  thick: '3px solid',
  dashed: '1px dashed',
  dotted: '1px dotted',
  double: '3px double',
};

/** Inline CSS approximating how the grid would draw a cell in this style. */
function previewCss(s: CellStyle): CSSProperties {
  const edge = (e: CellStyle['border']['top']) =>
    e.style === 'none'
      ? '1px dashed var(--cs-border)'
      : `${CSS_BORDER[e.style] ?? '1px solid'} ${Argb.toCss(e.color)}`;
  const decoration = [s.font.underline && 'underline', s.font.strikethrough && 'line-through'].filter(
    Boolean,
  );
  return {
    fontFamily: s.font.family,
    fontSize: `${s.font.sizePoints}pt`,
    fontWeight: s.font.bold ? 'bold' : 'normal',
    fontStyle: s.font.italic ? 'italic' : 'normal',
    textDecoration: decoration.length ? decoration.join(' ') : 'none',
    color: Argb.toCss(Argb.isTransparent(s.font.color) ? Argb.black : s.font.color),
    background: s.fill.background !== null ? Argb.toCss(s.fill.background) : 'transparent',
    justifyContent: { general: 'flex-end', left: 'flex-start', center: 'center', right: 'flex-end' }[
      s.alignment.horizontal
    ],
    alignItems: { top: 'flex-start', middle: 'center', bottom: 'flex-end' }[s.alignment.vertical],
    whiteSpace: s.alignment.wrapText ? 'normal' : 'nowrap',
    borderTop: edge(s.border.top),
    borderBottom: edge(s.border.bottom),
    borderLeft: edge(s.border.left),
    borderRight: edge(s.border.right),
  };
}
