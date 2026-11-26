import { useState } from 'preact/hooks';
import { Icon, region, ROVING_GROUP, rovingItem, rovingStep, type CommandCatalog } from '@cellestra/ui';
import type { AppController } from './controller.ts';

/** The classic toolbar (without the ribbon): one row of buttons, all from the catalog. */
const TOOLBAR: readonly (string | '|')[] = [
  'file.new',
  'file.open',
  'file.save',
  '|',
  'edit.undo',
  'edit.redo',
  '|',
  'format.bold',
  'format.italic',
  'format.underline',
  'format.strikethrough',
  '|',
  'format.alignLeft',
  'format.alignCenter',
  'format.alignRight',
  'format.wrap',
  'format.merge',
  '|',
  'format.asCurrency',
  'format.asPercent',
  'format.decreaseDecimal',
  'format.increaseDecimal',
  '|',
  'format.borderAll',
  'edit.formatPainter',
  'edit.autoSum',
  'data.sortAsc',
  'data.sortDesc',
  '|',
  'view.freezePanes',
  '|',
  'data.toggleFilter',
  'data.createTable',
  'insert.chartSheet',
  'insert.function',
  '|',
  'insert.rowsAbove',
  'insert.columnsLeft',
  '|',
  'edit.find',
  'edit.clearFormats',
  'format.cells',
];

/** The Command Menu button's place among the toolbar's stops. */
const PALETTE = 'palette';

/**
 * A toolbar is one Tab stop: ← → Home End move along its buttons (PRD §6.4), and the current one is
 * marked and focused from state.
 */
export function Toolbar({
  app,
  catalog,
  titleInline,
}: {
  readonly app: AppController;
  readonly catalog: CommandCatalog;
  /** Show the workbook's name in the toolbar (when there is no title bar of our own). */
  readonly titleInline: boolean;
}) {
  const s = app.session;
  // One Tab stop: the current button, kept by id so it stays put as others turn on and off.
  const [current, setCurrent] = useState<string | null>(null);
  // Read the signals the enabled/active checks depend on, so the toolbar re-renders with them.
  void s.version.value;
  void app.selection.state.value;
  void app.formatPainter.value;
  void app.recents.value;
  const stops = [
    ...TOOLBAR.filter((id) => {
      const c = id === '|' ? undefined : catalog.get(id);
      return c !== undefined && catalog.isEnabled(c);
    }),
    PALETTE,
  ];
  const at = current !== null && stops.includes(current) ? current : stops[0]!;
  return (
    <div
      class="cs-toolbar"
      role="toolbar"
      aria-label="Commands"
      {...ROVING_GROUP}
      {...region('toolbar')}
      onKeyDown={(e) => {
        const next = rovingStep(e, stops.indexOf(at), stops.length, 'horizontal');
        if (next === undefined) return;
        e.preventDefault();
        setCurrent(stops[next]!);
      }}
    >
      {/* Buttons that don't fit wrap onto a hidden second line: whole buttons drop off the
          end instead of the row overflowing. All of them are in ⌘K and the menus. */}
      <div class="cs-toolbar-items">
        {TOOLBAR.map((id, i) => {
          if (id === '|') return <span key={i} class="cs-sep" />;
          const c = catalog.get(id);
          if (!c) return null;
          const label = c.shortcutText ? `${c.title} (${c.shortcutText})` : c.title;
          return (
            <button
              key={id}
              type="button"
              class="cs-btn"
              title={label}
              aria-label={c.title}
              aria-pressed={c.active ? c.active() : undefined}
              disabled={!catalog.isEnabled(c)}
              {...rovingItem(id === at)}
              onPointerDown={(e) => e.preventDefault() /* keep focus (and the selection) in the grid */}
              onClick={() => void catalog.run(id)}
            >
              <Icon name={c.icon ?? 'square'} />
            </button>
          );
        })}
      </div>
      {titleInline ? (
        <span class="cs-title">
          {s.modified.value ? '• ' : ''}
          {s.title.value}
        </span>
      ) : (
        <span class="cs-toolbar-fill" />
      )}
      <button
        type="button"
        class="cs-btn"
        title="Command Menu (⌘K)"
        aria-label="Command Menu"
        {...rovingItem(at === PALETTE)}
        onClick={() => (app.paletteOpen.value = true)}
      >
        <Icon name="command" />
      </button>
    </div>
  );
}
