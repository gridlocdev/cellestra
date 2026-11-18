import {
  after,
  toAccelerator,
  type CommandCatalog,
  type CommandGroup,
  type NativeMenuItem,
  type Shortcut,
} from '@cellestra/ui';
import { effect } from '@preact/signals';
import type { AppController } from './controller.ts';

const MENU_ORDER: readonly CommandGroup[] = [
  'File',
  'Edit',
  'View',
  'Insert',
  'Format',
  'Data',
  'Sheet',
  'Help',
];
/** Clipboard items use Electron roles so the page gets real copy/cut/paste events. */
const ROLES: Record<string, string> = { 'edit.cut': 'cut', 'edit.copy': 'copy', 'edit.paste': 'paste' };
/** Web-only commands. */
const WEB_ONLY = new Set(['file.openLocal', 'file.download']);
/** Separators before these ids. */
const BREAKS = new Set([
  'file.save',
  'file.importCsv',
  'edit.cut',
  'edit.delete',
  'edit.find',
  'edit.autoSum',
  'edit.deleteRows',
  'view.zoomIn',
  'view.theme',
  'format.alignLeft',
  'format.asAutomatic',
  'format.borderAll',
  'format.columnWidth',
  'format.cleanUpSheet',
  'format.cells',
  'data.customSort',
  'data.toggleFilter',
  'data.createTable',
  'data.validation',
  'data.chooseFromList',
  'insert.function',
  'insert.date',
  'formula.recalc',
  'sheet.next',
  'sheet.resetData',
]);

/**
 * A chord without ⌘/Ctrl on a printable key (⇧Space, ⌥=) also types text. As a menu accelerator
 * the native menu would take it from every text field, so the page's own key handling keeps it.
 */
function typesText(s: Shortcut): boolean {
  return !s.mod && !s.ctrl && (s.key.length === 1 || s.key === 'space');
}

/** The desktop menu bar: the catalog in menu order, with each command's current state. */
export function nativeMenu(catalog: CommandCatalog): NativeMenuItem[] {
  return MENU_ORDER.map((group) => {
    const items: NativeMenuItem[] = [];
    for (const c of catalog.all) {
      if (c.group !== group || WEB_ONLY.has(c.id)) continue;
      if (BREAKS.has(c.id) && items.length > 0) items.push({ type: 'separator' });
      const role = ROLES[c.id];
      if (role) {
        items.push({ role });
        continue;
      }
      items.push({
        label: c.title,
        command: c.id,
        enabled: catalog.isEnabled(c),
        ...(c.shortcutKey && !typesText(c.shortcutKey) ? { accelerator: toAccelerator(c.shortcutKey) } : {}),
        ...(c.active ? { type: 'checkbox' as const, checked: c.active() } : {}),
      });
    }
    return { label: group, submenu: items };
  });
}

/** Sends the menu to the desktop app, dropping one identical to the last one sent. */
export interface MenuSender {
  /** True when the menu was sent. */
  send(menu: readonly NativeMenuItem[]): boolean;
  /** Send the next menu even if it is unchanged (another window took the menu bar meanwhile). */
  forget(): void;
}

/**
 * Electron rebuilds the whole menu bar for each menu it gets (and closes a menu that is open), so
 * moving the cursor around cells that enable the same commands must not reach it.
 */
export function menuSender(send: (menu: readonly NativeMenuItem[]) => void): MenuSender {
  let last = '';
  return {
    send(menu) {
      const key = JSON.stringify(menu);
      if (key === last) return false;
      last = key;
      send(menu);
      return true;
    },
    forget() {
      last = '';
    },
  };
}

/**
 * When to work the menu out again. An edit (or a theme, recents or view change) is sent on the
 * next tick, so its shortcuts (⌘Z after typing) are live at once. A cursor move waits until the
 * cursor has stood still for `settleMs`, so holding an arrow key doesn't check every command on
 * every step. `refresh` builds and sends the menu.
 */
export function menuSchedule(refresh: () => void, settleMs: number): (soon: boolean) => void {
  let cancel: (() => void) | undefined;
  let urgent = false;
  return (soon) => {
    urgent ||= soon;
    cancel?.();
    cancel = after(urgent ? 0 : settleMs, () => {
      urgent = false;
      refresh();
    });
  };
}

/**
 * Call `changed` whenever something a menu item's enabled or checked state reads may have moved:
 * `changed(true)` for the document and the app's settings, `changed(false)` for the cursor, which
 * moves often and can wait for it to stand still. Returns a function that stops watching.
 */
export function watchMenuInputs(app: AppController, changed: (soon: boolean) => void): () => void {
  const stopApp = effect(() => {
    void app.session.version.value;
    void app.recents.value;
    void app.showFormulas.value;
    void app.isDark.value;
    void app.layout.value;
    void app.ribbonCollapsed.value;
    void app.formatPainter.value;
    void app.findQuery.value;
    void app.nasAvailable.value;
    changed(true);
  });
  const stopCursor = effect(() => {
    void app.selection.state.value;
    changed(false);
  });
  return () => {
    stopApp();
    stopCursor();
  };
}
