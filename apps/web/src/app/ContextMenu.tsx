import { useEffect, useLayoutEffect, useRef, useState } from 'preact/hooks';
import { Icon, ROVING_GROUP, rovingItem, rovingStep, takeFocus, type CommandCatalog } from '@cellestra/ui';
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
    'edit.moveColumnsLeft',
    'edit.moveColumnsRight',
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
    'edit.moveRowsUp',
    'edit.moveRowsDown',
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
  // The item the arrow keys are on, among the enabled ones; null until an arrow key is pressed, so a
  // right-click leaves the keyboard on the grid.
  const [current, setCurrent] = useState<number | null>(null);
  const pickable = useRef<readonly string[]>([]);

  const close = () => {
    app.contextMenu.value = null;
    app.grid.current?.focus();
  };

  useEffect(() => {
    setCurrent(null);
    if (!menu) return;
    const away = (e: PointerEvent) => {
      if (!ref.current?.contains(e.target as Node)) close();
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        close();
        return;
      }
      const count = pickable.current.length;
      // From no item yet, ↑ starts at the bottom and ↓ at the top.
      const from = (c: number | null) => c ?? (e.key === 'ArrowUp' ? 0 : -1);
      if (rovingStep(e, from(null), count, 'vertical') === undefined) return;
      e.preventDefault();
      e.stopPropagation();
      setCurrent((c) => rovingStep(e, from(c), count, 'vertical') ?? c);
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
  const { row, column } = app.selection.active.peek();
  const list =
    menu.zone === 'cell' && app.sheet.findValidation(row, column)?.kind === 'list'
      ? ['data.chooseFromList', '|']
      : [];
  const items =
    menu.zone === 'cell' && app.activeTable
      ? [...list, ...MENUS.cell, ...TABLE_ITEMS]
      : [...list, ...MENUS[menu.zone]];
  pickable.current = items.filter((id) => {
    const c = id === '|' ? undefined : catalog.get(id);
    return c !== undefined && !c.hidden && catalog.isEnabled(c);
  });
  const focused = current === null ? undefined : pickable.current[current];
  return (
    <div
      ref={ref}
      class="cs-context-menu"
      role="menu"
      aria-label="Cell actions"
      {...ROVING_GROUP}
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
            {...rovingItem(id === focused, takeFocus)}
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
