import { effect } from '@preact/signals';
import { render } from 'preact';
import { toAccelerator, type CommandCatalog, type CommandGroup, type NativeMenuItem } from '@cellestra/ui';
import '@cellestra/ui/styles.css';
import { App } from './app/App.tsx';
import { buildCommands } from './app/commands.ts';
import { AppController } from './app/controller.ts';

const app = new AppController();
const catalog = buildCommands(app);

render(<App app={app} catalog={catalog} />, document.getElementById('app')!);
void app.start();
// Dev-only handle for poking at the app from the console.
if (import.meta.env.DEV) Object.assign(globalThis, { cellestra: { app, catalog } });

// ── Electron ────────────────────────────────────────────────────────────
const native = app.native;
if (native) {
  // Menu accelerators fire before the page sees the key. In a text field, grid commands give the
  // field its usual behaviour instead (⌘A selects the text, ⌘Z undoes typing).
  const TEXT_FALLBACK: Record<string, string> = {
    'edit.selectAll': 'selectAll',
    'edit.undo': 'undo',
    'edit.redo': 'redo',
  };
  native.onCommand((id) => {
    const inText = !!document.activeElement?.closest('input, textarea, select, [contenteditable]');
    if (inText && !catalog.get(id)?.global) {
      const fallback = TEXT_FALLBACK[id];
      if (fallback) document.execCommand(fallback);
      return;
    }
    void catalog.run(id);
  });
  native.onOpenFile((path) => void app.openNativePath(path));
  native.onCloseRequested(() => void app.requestClose());
  native.onExternalChange(() => void app.externalChange());

  effect(() => {
    const ref = app.session.document.value;
    native.setDocumentState({
      title: app.session.title.value,
      dirty: app.session.modified.value,
      path: ref?.store === 'native' ? ref.id : null,
    });
  });

  // Rebuild the menu when enabled/checked states may have moved; coalesced to one send per frame.
  let queued = false;
  effect(() => {
    void app.session.version.value;
    void app.selection.state.value;
    void app.recents.value;
    void app.showFormulas.value;
    void app.isDark.value;
    if (queued) return;
    queued = true;
    requestAnimationFrame(() => {
      queued = false;
      native.setMenu(nativeMenu(catalog));
    });
  });
}

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
  'formula.recalc',
  'sheet.next',
]);

function nativeMenu(catalog: CommandCatalog): NativeMenuItem[] {
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
        ...(c.shortcutKey ? { accelerator: toAccelerator(c.shortcutKey) } : {}),
        ...(c.active ? { type: 'checkbox' as const, checked: c.active() } : {}),
      });
    }
    return { label: group, submenu: items };
  });
}
