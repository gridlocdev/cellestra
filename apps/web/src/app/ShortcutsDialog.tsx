import { useCallback, useMemo } from 'preact/hooks';
import {
  Dialog,
  KeyLookup,
  type CellestraNative,
  type Command,
  type CommandCatalog,
  type KeyCapture,
  type KeyLookupResult,
} from '@cellestra/ui';

/** The desktop app's key capture, when this build has one (PRD §6.4). */
function keyCaptureOf(native: CellestraNative | null): KeyCapture | null {
  const capture = native?.captureKeys?.bind(native);
  const onCaptured = native?.onCapturedKey?.bind(native);
  if (!capture || !onCaptured) return null;
  return { captureKeys: capture, onCapturedKey: onCaptured };
}

/**
 * Every shortcut, grouped as in the menus, and "Look Up a Key": press keys to see which command
 * each one runs. While looking up, keys run nothing. Escape, Stop, leaving the window, closing the
 * dialog or 30 seconds without a key all end it.
 */
export function ShortcutsDialog({
  catalog,
  native,
  onClose,
}: {
  readonly catalog: CommandCatalog;
  readonly native: CellestraNative | null;
  readonly onClose: () => void;
}) {
  const lookup = useMemo(
    () => new KeyLookup(catalog, { target: window, capture: keyCaptureOf(native) }),
    [catalog, native],
  );
  // Closing or unmounting the dialog always ends the lookup.
  const ended = useCallback(
    (el: HTMLElement | null) => {
      if (!el) lookup.stop();
    },
    [lookup],
  );
  const close = () => {
    lookup.stop();
    onClose();
  };
  const listening = lookup.listening.value;
  const last = lookup.last.value;
  const found = new Set(last?.commands.map((c) => c.id) ?? []);

  const groups = new Map<string, Command[]>();
  for (const c of catalog.all) {
    if (!c.shortcutText || c.hidden) continue;
    const list = groups.get(c.group) ?? [];
    list.push(c);
    groups.set(c.group, list);
  }
  return (
    <Dialog
      title="Keyboard Shortcuts"
      onClose={close}
      actions={
        <button type="submit" class="cs-btn primary">
          Done
        </button>
      }
      onSubmit={close}
    >
      <div ref={ended} class="cs-key-lookup" aria-live="polite" style={LOOKUP_STYLE}>
        <button type="button" class="cs-btn" onClick={() => (listening ? lookup.stop() : lookup.start())}>
          {listening ? 'Stop' : 'Look Up a Key…'}
        </button>
        <span style={{ fontSize: '13px' }}>
          <LookupText listening={listening} last={last} native={native !== null} />
        </span>
      </div>
      <div style={{ columns: '2', columnGap: '28px', minWidth: '560px' }}>
        {[...groups].map(([group, items]) => (
          <section key={group} style={{ breakInside: 'avoid', marginBottom: '12px' }}>
            <h3 style={{ margin: '0 0 4px', fontSize: '13px' }}>{group}</h3>
            {items.map((c) => (
              <div
                key={c.id}
                data-command={c.id}
                ref={found.has(c.id) ? showFound : undefined}
                class={found.has(c.id) ? 'cs-shortcut found' : 'cs-shortcut'}
                style={{
                  display: 'flex',
                  gap: '12px',
                  fontSize: '13px',
                  lineHeight: '22px',
                  padding: '0 4px',
                  borderRadius: '4px',
                  scrollMarginTop: '48px',
                  background: found.has(c.id) ? 'var(--cs-hover)' : undefined,
                  fontWeight: found.has(c.id) ? 600 : undefined,
                }}
              >
                <span>{c.title}</span>
                <span class="cs-kbd" style={{ marginLeft: 'auto', opacity: 0.7 }}>
                  {c.shortcutText}
                </span>
              </div>
            ))}
          </section>
        ))}
      </div>
    </Dialog>
  );
}

/** A found row further down the list scrolls into view. */
const showFound = (el: HTMLElement | null) => el?.scrollIntoView({ block: 'nearest' });

const LOOKUP_STYLE = {
  display: 'flex',
  alignItems: 'center',
  gap: '12px',
  margin: '-8px 0 6px',
  padding: '8px 0',
  minHeight: '28px',
  // Stays in view while the list scrolls to a found row.
  position: 'sticky',
  top: '-8px',
  zIndex: 1,
  background: 'var(--cs-bg)',
} as const;

function LookupText({
  listening,
  last,
  native,
}: {
  readonly listening: boolean;
  readonly last: KeyLookupResult | null;
  readonly native: boolean;
}) {
  if (last) {
    const what =
      last.commands.length > 0
        ? last.commands.map((c) => `${c.title} (${c.group})`).join(', ')
        : last.gridKey
          ? 'no command; the grid uses it to type, move or edit'
          : 'no command';
    return (
      <>
        <span class="cs-kbd">{last.keys}</span>: {what}
        {listening ? '. Press another, or Esc to stop.' : ''}
      </>
    );
  }
  if (listening) return <>Press any keys to see what they do. Esc stops.</>;
  return native ? (
    <>Press keys to see which command each one runs.</>
  ) : (
    <>Press keys to see which command each one runs. The browser keeps some keys, such as ⌘W and ⌘T.</>
  );
}
