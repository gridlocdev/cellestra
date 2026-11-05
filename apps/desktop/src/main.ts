import { app, BrowserWindow, dialog, ipcMain, Menu, net, protocol, shell } from 'electron';
import type { MenuItemConstructorOptions } from 'electron';
import { watch, type FSWatcher } from 'node:fs';
import { open, readFile, rename, stat, unlink } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import {
  Channels,
  CONFLICT,
  type DocumentState,
  type NativeFile,
  type NativeMenuItem,
  type SavedFile,
} from './bridge.ts';

const here = path.dirname(fileURLToPath(import.meta.url));
const webRoot = path.join(here, 'web');
const FILTERS = [
  { name: 'Spreadsheets', extensions: ['ods', 'csv', 'tsv'] },
  { name: 'OpenDocument Spreadsheet', extensions: ['ods'] },
];

// The web client loads from app://cellestra/ — a standard, secure scheme so fetch, modules and
// CSP behave like https, without file:// quirks.
protocol.registerSchemesAsPrivileged([
  { scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true } },
]);

/** Per-window state the main process must know: allow-close flag, file watcher. */
interface WindowState {
  allowClose: boolean;
  watcher: FSWatcher | null;
  watchedPath: string | null;
  /** Debounce for watcher events; cleared with the watcher so it never fires into a closed window. */
  watchTimer: NodeJS.Timeout | null;
}
const windows = new Map<number, WindowState>();
/** Files asked for before a window was ready (macOS open-file at launch, argv). */
const pendingOpen: string[] = process.argv
  .slice(app.isPackaged ? 1 : 2)
  .filter((a) => /\.(ods|csv|tsv)$/i.test(a));

function stopWatching(state: WindowState): void {
  state.watcher?.close();
  state.watcher = null;
  if (state.watchTimer) clearTimeout(state.watchTimer);
  state.watchTimer = null;
}

function createWindow(): BrowserWindow {
  const win = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 640,
    minHeight: 400,
    title: 'Cellestra',
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    backgroundColor: '#1e1e1e',
    show: false,
    webPreferences: {
      preload: path.join(here, 'preload.cjs'),
      contextIsolation: true,
      sandbox: true,
      nodeIntegration: false,
      spellcheck: false,
    },
  });
  const state: WindowState = { allowClose: false, watcher: null, watchedPath: null, watchTimer: null };
  windows.set(win.id, state);

  win.once('ready-to-show', () => win.show());
  win.on('close', (e) => {
    if (state.allowClose) return;
    e.preventDefault();
    win.webContents.send(Channels.closeRequested);
  });
  win.on('closed', () => {
    stopWatching(state);
    windows.delete(win.id);
  });
  // Links (cell hyperlinks) open in the user's browser, never inside the app.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^(https?|mailto):/i.test(url)) void shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e, url) => {
    if (!url.startsWith('app://')) e.preventDefault();
  });
  win.webContents.on('did-finish-load', () => {
    for (const p of pendingOpen.splice(0)) win.webContents.send(Channels.openFile, p);
  });

  void win.loadURL('app://cellestra/index.html');
  return win;
}

function fromEvent(e: Electron.IpcMainEvent | Electron.IpcMainInvokeEvent): BrowserWindow | null {
  return BrowserWindow.fromWebContents(e.sender);
}

async function readNative(p: string): Promise<NativeFile> {
  const [bytes, s] = await Promise.all([readFile(p), stat(p)]);
  app.addRecentDocument(p);
  return { path: p, name: path.basename(p), bytes: new Uint8Array(bytes), modified: s.mtimeMs };
}

/** Atomic save: temp file beside the target, fsync, rename. Refuses if the file changed since opened. */
async function writeNative(
  p: string,
  bytes: Uint8Array,
  expectedModified: number | null,
  force: boolean,
): Promise<SavedFile> {
  if (!force && expectedModified !== null) {
    const s = await stat(p).catch(() => null);
    if (s && Math.abs(s.mtimeMs - expectedModified) > 1)
      throw new Error(`${CONFLICT}: the file changed on disk`);
  }
  const tmp = path.join(path.dirname(p), `.${path.basename(p)}.${process.pid}.${Date.now()}.tmp`);
  try {
    const h = await open(tmp, 'wx', 0o644);
    try {
      await h.writeFile(bytes);
      await h.sync();
    } finally {
      await h.close();
    }
    await rename(tmp, p);
  } catch (e) {
    await unlink(tmp).catch(() => {});
    throw e;
  }
  const s = await stat(p);
  app.addRecentDocument(p);
  return { path: p, name: path.basename(p), modified: s.mtimeMs };
}

function toTemplate(items: readonly NativeMenuItem[], win: BrowserWindow): MenuItemConstructorOptions[] {
  return items.map((item): MenuItemConstructorOptions => {
    if (item.type === 'separator') return { type: 'separator' };
    if (item.role)
      return {
        role: item.role as MenuItemConstructorOptions['role'],
        ...(item.label ? { label: item.label } : {}),
      };
    const opts: MenuItemConstructorOptions = {
      label: item.label ?? '',
      enabled: item.enabled ?? true,
      ...(item.accelerator ? { accelerator: item.accelerator } : {}),
      ...(item.type === 'checkbox' ? { type: 'checkbox', checked: item.checked ?? false } : {}),
    };
    if (item.submenu) opts.submenu = toTemplate(item.submenu, win);
    else if (item.command) {
      const id = item.command;
      opts.click = () => win.webContents.send(Channels.command, id);
    }
    return opts;
  });
}

function registerIpc(): void {
  ipcMain.handle(Channels.openDialog, async (e) => {
    const win = fromEvent(e);
    const result = win
      ? await dialog.showOpenDialog(win, { properties: ['openFile'], filters: FILTERS })
      : await dialog.showOpenDialog({ properties: ['openFile'], filters: FILTERS });
    const p = result.filePaths[0];
    return result.canceled || !p ? null : readNative(p);
  });
  ipcMain.handle(Channels.readFile, (_e, p: string) => readNative(p));
  ipcMain.handle(
    Channels.saveFile,
    (_e, p: string, bytes: Uint8Array, expected: number | null, force: boolean) =>
      writeNative(p, bytes, expected, force),
  );
  ipcMain.handle(Channels.saveDialog, async (e, suggestedName: string, bytes: Uint8Array) => {
    const win = fromEvent(e);
    const opts = { defaultPath: suggestedName, filters: FILTERS };
    const result = win ? await dialog.showSaveDialog(win, opts) : await dialog.showSaveDialog(opts);
    return result.canceled || !result.filePath ? null : writeNative(result.filePath, bytes, null, true);
  });
  ipcMain.on(Channels.setDocumentState, (e, s: DocumentState) => {
    const win = fromEvent(e);
    if (!win) return;
    win.setTitle(s.title);
    win.setDocumentEdited(s.dirty);
    if (process.platform === 'darwin') win.setRepresentedFilename(s.path ?? '');
  });
  ipcMain.on(Channels.setMenu, (e, items: readonly NativeMenuItem[]) => {
    const win = fromEvent(e);
    if (!win) return;
    const template = toTemplate(items, win);
    if (process.platform === 'darwin') template.unshift({ role: 'appMenu' });
    Menu.setApplicationMenu(Menu.buildFromTemplate(template));
  });
  ipcMain.on(Channels.confirmClose, (e) => {
    const win = fromEvent(e);
    const state = win && windows.get(win.id);
    if (!win || !state) return;
    state.allowClose = true;
    win.close();
  });
  ipcMain.on(Channels.watch, (e, p: string | null) => {
    const win = fromEvent(e);
    const state = win && windows.get(win.id);
    if (!win || !state) return;
    stopWatching(state);
    state.watchedPath = p;
    if (!p) return;
    try {
      state.watcher = watch(p, () => {
        // Atomic saves arrive as several events; debounce, and let the renderer compare mtimes.
        if (state.watchTimer) clearTimeout(state.watchTimer);
        state.watchTimer = setTimeout(() => {
          state.watchTimer = null;
          if (!win.isDestroyed()) win.webContents.send(Channels.externalChange, p);
        }, 300);
      });
    } catch {
      state.watcher = null;
    }
  });
}

app.on('open-file', (e, p) => {
  e.preventDefault();
  const win = BrowserWindow.getAllWindows()[0];
  if (win && !win.webContents.isLoading()) win.webContents.send(Channels.openFile, p);
  else pendingOpen.push(p);
});

const single = app.requestSingleInstanceLock();
if (!single) app.quit();
app.on('second-instance', (_e, argv) => {
  const win = BrowserWindow.getAllWindows()[0];
  if (!win) return;
  if (win.isMinimized()) win.restore();
  win.focus();
  for (const a of argv.filter((x) => /\.(ods|csv|tsv)$/i.test(x))) win.webContents.send(Channels.openFile, a);
});

void app.whenReady().then(() => {
  protocol.handle('app', (req) => {
    const { pathname } = new URL(req.url);
    const file = path.normalize(path.join(webRoot, decodeURIComponent(pathname)));
    if (!file.startsWith(webRoot + path.sep)) return new Response('Forbidden', { status: 403 });
    return net.fetch(pathToFileURL(file).toString());
  });
  registerIpc();
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
