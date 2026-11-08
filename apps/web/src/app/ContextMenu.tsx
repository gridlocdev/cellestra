import { useEffect, useLayoutEffect, useRef, useState } from 'preact/hooks';
import { Icon, type CommandCatalog } from '@cellestra/ui';
import type { AppController } from './controller.ts';

/** What each part of the grid offers on right-click, as command ids (`|` is a separator). */
const MENUS: Record<'cell' | 'columnHeader' | 'rowHeader', readonly string[]> = {
  cell: [
    'edit.cut',
    'edit.copy',
    'edit.paste',
    '|',
    'insert.rowsAbove',
    'insert.columnsLeft',
    '|',
    'edit.delete',
    'edit.clearFormats',
    'edit.clearAll',
    '|',
    'format.merge',
    'format.cells',
  ],
  columnHeader: [
    'edit.cut',
    'edit.copy',
    'edit.paste',
    '|',
    'insert.columnsLeft',
    'edit.deleteColumns',
    '|',
    'format.columnWidth',
    'data.autofitColumns',
    'format.hideColumns',
    'format.unhideColumns',
    '|',
    'edit.delete',
    'edit.clearFormats',
    '|',
    'format.cells',
  ],
  rowHeader: [
    'edit.cut',
    'edit.copy',
    'edit.paste',
    '|',
    'insert.rowsAbove',
    'edit.deleteRows',
    '|',
    'format.rowHeight',
    'format.hideRows',
    'format.unhideRows',
    '|',
    'edit.delete',
    'edit.clearFormats',
    '|',
    'format.cells',
  ],
};

/** Added to the cell menu when the click lands in a table: its own formatting, ticked when on. */
const TABLE_ITEMS: readonly string[] = ['|', 'data.bandedRows', 'data.headerColumn', 'data.removeTable'];

/** The grid's right-click menu: the catalog's own commands, so titles and shortcuts match the rest. */
export function GridContextMenu({
  app,
  catalog,
}: {
  readonly app: AppController;
  readonly catalog: CommandCatalog;
}) {
  const menu = app.contextMenu.value;
  const ref = useRef<HTMLDivElement>(null);
  const [at, setAt] = useState<{ x: number; y: number } | null>(null);

  const close = () => {
    app.contextMenu.value = null;
    app.grid.current?.focus();
  };

  useEffect(() => {
    if (!menu) return;
    const away = (e: PointerEvent) => {
      if (!ref.current?.contains(e.target as Node)) close();
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        close();
      }
    };
    addEventListener('pointerdown', away, true);
    addEventListener('keydown', key, true);
    addEventListener('blur', close);
    return () => {
      removeEventListener('pointerdown', away, true);
      removeEventListener('keydown', key, true);
      removeEventListener('blur', close);
    };
  }, [menu]);

  // Keep the menu on screen: flip it up or left when it would run off the window.
  useLayoutEffect(() => {
    if (!menu) return setAt(null);
    const el = ref.current;
    const w = el?.offsetWidth ?? 220;
    const h = el?.offsetHeight ?? 300;
    setAt({
      x: menu.x + w > innerWidth - 4 ? Math.max(4, menu.x - w) : menu.x,
      y: menu.y + h > innerHeight - 4 ? Math.max(4, menu.y - h) : menu.y,
    });
  }, [menu]);

  if (!menu) return null;
  const items = menu.zone === 'cell' && app.activeTable ? [...MENUS.cell, ...TABLE_ITEMS] : MENUS[menu.zone];
  return (
    <div
      ref={ref}
      class="cs-context-menu"
      role="menu"
      aria-label="Cell actions"
      style={{
        left: `${(at ?? menu).x}px`,
        top: `${(at ?? menu).y}px`,
        visibility: at ? 'visible' : 'hidden',
      }}
      onContextMenu={(e) => e.preventDefault()}
    >
      {items.map((id, i) => {
        if (id === '|') return <div key={i} class="cs-context-sep" role="separator" />;
        const c = catalog.get(id);
        if (!c || c.hidden) return null;
        const on = c.active?.();
        return (
          <button
            key={id}
            type="button"
            role={on === undefined ? 'menuitem' : 'menuitemcheckbox'}
            aria-checked={on}
            class="cs-context-item"
            disabled={!catalog.isEnabled(c)}
            onPointerDown={(e) => e.preventDefault()}
            onClick={() => {
              close();
              catalog.run(id);
            }}
          >
            <Icon name={c.icon ?? 'square'} />
            <span class="cs-context-title">{c.title}</span>
            {c.shortcutText && <span class="cs-context-key">{c.shortcutText}</span>}
            {on !== undefined && <span class="cs-context-check">{on ? <Icon name="check" /> : null}</span>}
          </button>
        );
      })}
    </div>
  );
}
