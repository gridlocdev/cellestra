import { effect } from '@preact/signals';
import { render } from 'preact';
import '@cellestra/ui/styles.css';
import { App } from './app/App.tsx';
import { buildCommands } from './app/commands.ts';
import { AppController } from './app/controller.ts';
import { menuSchedule, menuSender, nativeMenu, watchMenuInputs } from './app/nativeMenu.ts';

/** How long the cursor must stand still before the desktop menu is worked out again. */
const MENU_SETTLE_MS = 100;

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

  // Enabled/checked states may have moved: work the menu out again, and send it only when it
  // differs from what the app shows.
  const menu = menuSender((items) => native.setMenu(items));
  const menuChanged = menuSchedule(() => menu.send(nativeMenu(catalog)), MENU_SETTLE_MS);
  // The menu bar is shared: the window that sent last owns it, so a window coming to the front
  // claims it back even when its own menu didn't change.
  addEventListener('focus', () => {
    menu.forget();
    menuChanged(true);
  });
  watchMenuInputs(app, menuChanged);
}
