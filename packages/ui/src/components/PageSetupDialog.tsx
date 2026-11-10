import { PAPER_SIZES, PageSetups, type PageSetup } from '@cellestra/core';
import { useMemo, useState } from 'preact/hooks';
import { Dialog } from './Dialog.tsx';
import { draftOf, setupOf, type PageSetupDraft } from './pageSetupForm.ts';

/**
 * How the active sheet prints: paper, orientation, margins, scale, grid lines and headers, centring
 * and page order. Saved in the file, so LibreOffice prints the sheet the same way.
 */
export function PageSetupDialog(props: {
  readonly sheetName: string;
  readonly current: PageSetup | null;
  readonly onApply: (setup: PageSetup) => void;
  /** Apply, then print with the new setup. */
  readonly onPrint?: (setup: PageSetup) => void;
  readonly onClose: () => void;
}) {
  const [draft, setDraft] = useState<PageSetupDraft>(() => draftOf(props.current));
  const [tried, setTried] = useState(false);
  const result = useMemo(() => setupOf(draft, props.current), [draft, props.current]);
  const error = 'error' in result ? result.error : null;
  const set = (patch: Partial<PageSetupDraft>) => setDraft((d) => ({ ...d, ...patch }));
  const custom = props.current && PageSetups.paperOf(props.current) === null ? props.current : null;

  const finish = (then?: (setup: PageSetup) => void) => {
    setTried(true);
    if ('error' in result) return;
    props.onApply(result.setup);
    props.onClose();
    then?.(result.setup);
  };

  const text = (key: keyof PageSetupDraft, label: string, extra: Record<string, unknown> = {}) => (
    <input
      aria-label={label}
      inputMode="decimal"
      value={draft[key] as string}
      onInput={(e) => set({ [key]: (e.target as HTMLInputElement).value } as Partial<PageSetupDraft>)}
      {...extra}
    />
  );
  const check = (key: keyof PageSetupDraft, label: string) => (
    <label class="cs-ps-check">
      <input
        type="checkbox"
        checked={draft[key] as boolean}
        onChange={(e) => set({ [key]: (e.target as HTMLInputElement).checked } as Partial<PageSetupDraft>)}
      />
      {label}
    </label>
  );
  const radio = <K extends keyof PageSetupDraft>(key: K, value: PageSetupDraft[K], label: string) => (
    <label class="cs-ps-check">
      <input
        type="radio"
        name={key}
        checked={draft[key] === value}
        onChange={() => set({ [key]: value } as Partial<PageSetupDraft>)}
      />
      {label}
    </label>
  );

  return (
    <Dialog
      title={`Page Setup — ${props.sheetName}`}
      class="cs-page-setup"
      onClose={props.onClose}
      onSubmit={() => finish()}
      actions={
        <>
          <button type="button" class="cs-btn" onClick={props.onClose}>
            Cancel
          </button>
          {props.onPrint && (
            <button type="button" class="cs-btn" onClick={() => finish(props.onPrint)}>
              Print…
            </button>
          )}
          <button type="submit" class="cs-btn primary">
            OK
          </button>
        </>
      }
    >
      <fieldset>
        <legend>Paper</legend>
        <div class="cs-ps-row">
          <select
            aria-label="Paper size"
            value={draft.paper}
            onChange={(e) => set({ paper: (e.target as HTMLSelectElement).value })}
          >
            {PAPER_SIZES.map((p) => (
              <option key={p.id} value={p.id}>
                {p.label} ({p.width} × {p.height} mm)
              </option>
            ))}
            {custom && (
              <option value="custom">
                The file's own ({Math.round(Math.min(custom.paperWidth, custom.paperHeight))} ×{' '}
                {Math.round(Math.max(custom.paperWidth, custom.paperHeight))} mm)
              </option>
            )}
          </select>
          {radio('orientation', 'portrait', 'Portrait')}
          {radio('orientation', 'landscape', 'Landscape')}
        </div>
      </fieldset>
      <fieldset>
        <legend>Margins (mm)</legend>
        <div class="cs-ps-margins">
          <label>Top {text('top', 'Top margin')}</label>
          <label>Bottom {text('bottom', 'Bottom margin')}</label>
          <label>Left {text('left', 'Left margin')}</label>
          <label>Right {text('right', 'Right margin')}</label>
        </div>
      </fieldset>
      <fieldset>
        <legend>Scale</legend>
        <div class="cs-ps-row">
          {radio('scaleKind', 'percent', 'Adjust to')}
          {text('percent', 'Scale percent', {
            class: 'cs-ps-small',
            disabled: draft.scaleKind !== 'percent',
          })}
          %
        </div>
        <div class="cs-ps-row">
          {radio('scaleKind', 'fit', 'Fit to')}
          {text('fitWidth', 'Pages wide', { class: 'cs-ps-small', disabled: draft.scaleKind !== 'fit' })}
          <span>wide by</span>
          {text('fitHeight', 'Pages tall', {
            class: 'cs-ps-small',
            placeholder: 'any',
            disabled: draft.scaleKind !== 'fit',
          })}
          <span>tall</span>
        </div>
        <div class="cs-ps-row">
          {radio('scaleKind', 'pages', 'Fit on')}
          {text('pages', 'Number of pages', { class: 'cs-ps-small', disabled: draft.scaleKind !== 'pages' })}
          <span>pages</span>
        </div>
      </fieldset>
      <fieldset>
        <legend>Print</legend>
        <div class="cs-ps-row">
          {check('gridlines', 'Grid lines')}
          {check('headers', 'Row and column headers')}
        </div>
        <div class="cs-ps-row">
          {check('centerHorizontally', 'Center horizontally')}
          {check('centerVertically', 'Center vertically')}
        </div>
        <div class="cs-ps-row">
          {radio('pageOrder', 'downThenOver', 'Down, then over')}
          {radio('pageOrder', 'overThenDown', 'Over, then down')}
        </div>
      </fieldset>
      {tried && error && (
        <p class="cs-ps-error" role="alert">
          {error}
        </p>
      )}
    </Dialog>
  );
}
