import { app, BrowserWindow, dialog, ipcMain, Menu, net, protocol, shell } from 'electron';
import type { MenuItemConstructorOptions } from 'electron';
import { readFile, stat } from 'node:fs/promises';
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
import {
  isAppUrl,
  isSpreadsheetPath,
  PathGrants,
  spreadsheetArgs,
  watchFile,
  withFileLock,
  writeAtomic,
} from './files.ts';

const here = path.dirname(fileURLToPath(import.meta.url));
const webRoot = path.join(here, 'web');
const FILTERS = [
  { name: 'Spreadsheets', extensions: ['ods', 'csv', 'tsv'] },
  { name: 'OpenDocument Spreadsheet', extensions: ['ods'] },
];
/** How long the renderer has to acknowledge a close request before we offer to force it. */
const CLOSE_ACK_MS = 3000;

// The web client loads from app://cellestra/ — a standard, secure scheme so fetch, modules and
// CSP behave like https, without file:// quirks.
protocol.registerSchemesAsPrivileged([
  { scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true } },
]);

/** Per-window state the main process must know: allow-close flag, file watcher, renderer health. */
interface WindowState {
  allowClose: boolean;
  stopWatch: (() => void) | null;
  /** Running while a close request waits for the renderer's ack. */
  closeTimer: NodeJS.Timeout | null;
  /** The renderer crashed or was killed: nothing is left to ask about unsaved changes. */
  rendererGone: boolean;
  unresponsive: boolean;
}
const windows = new Map<number, WindowState>();
/** Files the renderer may touch: picked in a dialog, opened by the OS, or opened before. */
const grants = new PathGrants(path.join(app.getPath('userData'), 'granted-files.json'));
/**
 * Another instance is already running: it gets our argv (second-instance) and owns
 * granted-files.json, so this one must not write it on its way out.
 */
const single = app.requestSingleInstanceLock();
if (!single) {
  grants.disablePersist();
  app.quit();
}
/** mtime we last read or wrote per path, so the watcher can ignore our own saves. */
const knownMtime = new Map<string, number>();
/** Set by Cmd+Q / app.quit() so the last window closing ends the app on macOS too. */
let quitting = false;

function grantAll(paths: readonly string[]): string[] {
  for (const p of paths) grants.grant(p);
  return [...paths];
}

/** Files asked for before a window was ready (macOS open-file at launch, argv). */
const pendingOpen: string[] = grantAll(
  spreadsheetArgs(process.argv.slice(app.isPackaged ? 1 : 2), process.cwd()),
);

function stopWatching(state: WindowState): void {
  state.stopWatch?.();
  state.stopWatch = null;
}

function createWindow(): BrowserWindow {
  const win = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 640,
    minHeight: 400,
    title: 'Cellestra',
    // 'hidden' keeps the controls at their usual spot, centred in the web app's 28px title bar.
    titleBarStyle: process.platform === 'darwin' ? 'hidden' : 'default',
    backgroundColor: '#1e1e1e',
    // macOS ignores this and shows the bundle's icon.icns.
    ...(process.platform === 'darwin' ? {} : { icon: path.join(here, '..', 'build', 'icon.png') }),
    show: false,
    webPreferences: {
      preload: path.join(here, 'preload.cjs'),
      contextIsolation: true,
      sandbox: true,
      nodeIntegration: false,
      spellcheck: false,
    },
  });
  const state: WindowState = {
    allowClose: false,
    stopWatch: null,
    closeTimer: null,
    rendererGone: false,
    unresponsive: false,
  };
  windows.set(win.id, state);

  win.once('ready-to-show', () => win.show());
  win.on('close', (e) => {
    // A dead renderer cannot answer; its last recovery snapshot is all there is to keep.
    if (state.allowClose || state.rendererGone) return;
    e.preventDefault();
    if (state.unresponsive) return void offerForceClose(win, state);
    if (state.closeTimer) return; // already asked; waiting for the renderer
    win.webContents.send(Channels.closeRequested);
    state.closeTimer = setTimeout(() => {
      state.closeTimer = null;
      void offerForceClose(win, state);
    }, CLOSE_ACK_MS);
  });
  win.on('closed', () => {
    stopWatching(state);
    if (state.closeTimer) clearTimeout(state.closeTimer);
    windows.delete(win.id);
  });
  win.on('unresponsive', () => (state.unresponsive = true));
  win.on('responsive', () => (state.unresponsive = false));
  win.webContents.on('render-process-gone', () => {
    state.rendererGone = true;
    stopWatching(state);
    // A close or quit was waiting on this renderer: finish it.
    if (state.closeTimer || quitting) {
      if (state.closeTimer) clearTimeout(state.closeTimer);
      state.closeTimer = null;
      win.destroy();
    }
  });
  // Links (cell hyperlinks) open in the user's browser, never inside the app.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^(https?|mailto):/i.test(url)) void shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e, url) => {
    if (!isAppUrl(url)) e.preventDefault();
  });
  win.webContents.on('did-finish-load', () => {
    for (const p of pendingOpen.splice(0)) win.webContents.send(Channels.openFile, p);
  });

  void win.loadURL('app://cellestra/index.html');
  return win;
}

/** The renderer is hung or never answered the close request: let the user force it. */
async function offerForceClose(win: BrowserWindow, state: WindowState): Promise<void> {
  if (win.isDestroyed()) return;
  const { response } = await dialog.showMessageBox(win, {
    type: 'warning',
    buttons: ['Close Anyway', 'Wait'],
    defaultId: 1,
    cancelId: 1,
    message: 'Cellestra is not responding.',
    detail: 'Closing now loses changes made since the last recovery snapshot.',
  });
  if (win.isDestroyed()) return;
  if (response === 0) {
    state.allowClose = true;
    win.destroy();
  } else quitting = false;
}

function fromEvent(e: Electron.IpcMainEvent | Electron.IpcMainInvokeEvent): BrowserWindow | null {
  return BrowserWindow.fromWebContents(e.sender);
}

/** Only our own page may use the bridge. */
function trusted(e: Electron.IpcMainEvent | Electron.IpcMainInvokeEvent): boolean {
  return isAppUrl(e.senderFrame?.url);
}

function forbidden(): never {
  throw new Error('Forbidden');
}

async function readNative(p: string): Promise<NativeFile> {
  const [bytes, s] = await Promise.all([readFile(p), stat(p)]);
  grants.grant(p);
  knownMtime.set(p, s.mtimeMs);
  app.addRecentDocument(p);
  return { path: p, name: path.basename(p), bytes: new Uint8Array(bytes), modified: s.mtimeMs };
}

/**
 * The renderer asks for a path we never handed it (an old recent, or a compromised page): open it
 * only if it exists and the user agrees.
 */
async function allowRead(win: BrowserWindow | null, p: unknown): Promise<string> {
  if (grants.has(p)) return p;
  if (!isSpreadsheetPath(p) || !path.isAbsolute(p)) forbidden();
  await stat(p); // ENOENT reaches the renderer as "not found"
  const opts = {
    type: 'question' as const,
    buttons: ['Open', 'Cancel'],
    defaultId: 0,
    cancelId: 1,
    message: `Open “${path.basename(p)}”?`,
    detail: p,
  };
  const { response } = win ? await dialog.showMessageBox(win, opts) : await dialog.showMessageBox(opts);
  if (response !== 0) throw new Error('Opening the file was cancelled');
  grants.grant(p);
  return p;
}

/** Atomic save that keeps mode and hard links. Refuses if the file changed since opened. */
async function writeNative(
  p: string,
  bytes: Uint8Array,
  expectedModified: number | null,
  force: boolean,
): Promise<SavedFile> {
  // One save per file at a time, so the mtime check and the write can't interleave with another save.
  return withFileLock(p, async () => {
    if (!force && expectedModified !== null) {
      const s = await stat(p).catch(() => null);
      if (s && Math.abs(s.mtimeMs - expectedModified) > 1)
        throw new Error(`${CONFLICT}: the file changed on disk`);
    }
    await writeAtomic(p, bytes);
    const s = await stat(p);
    grants.grant(p);
    knownMtime.set(p, s.mtimeMs);
    app.addRecentDocument(p);
    return { path: p, name: path.basename(p), modified: s.mtimeMs };
  });
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
    if (!trusted(e)) forbidden();
    const win = fromEvent(e);
    const result = win
      ? await dialog.showOpenDialog(win, { properties: ['openFile'], filters: FILTERS })
      : await dialog.showOpenDialog({ properties: ['openFile'], filters: FILTERS });
    const p = result.filePaths[0];
    return result.canceled || !p ? null : readNative(p);
  });
  ipcMain.handle(Channels.readFile, async (e, p: unknown) => {
    if (!trusted(e)) forbidden();
    return readNative(await allowRead(fromEvent(e), p));
  });
  ipcMain.handle(
    Channels.saveFile,
    (e, p: unknown, bytes: Uint8Array, expected: number | null, force: boolean) => {
      if (!trusted(e) || !grants.has(p)) forbidden();
      return writeNative(p, bytes, expected, force);
    },
  );
  ipcMain.handle(Channels.saveDialog, async (e, suggestedName: string, bytes: Uint8Array) => {
    if (!trusted(e)) forbidden();
    const win = fromEvent(e);
    const opts = { defaultPath: String(suggestedName), filters: FILTERS };
    const result = win ? await dialog.showSaveDialog(win, opts) : await dialog.showSaveDialog(opts);
    if (result.canceled || !result.filePath) return null;
    if (!isSpreadsheetPath(result.filePath)) throw new Error('Save as .ods, .csv, or .tsv');
    return writeNative(result.filePath, bytes, null, true);
  });
  ipcMain.on(Channels.setDocumentState, (e, s: DocumentState) => {
    const win = fromEvent(e);
    if (!win || !trusted(e)) return;
    win.setTitle(s.title);
    win.setDocumentEdited(s.dirty);
    if (process.platform === 'darwin') win.setRepresentedFilename(grants.has(s.path) ? s.path : '');
  });
  ipcMain.on(Channels.setTitleBarHeight, (e, height: unknown) => {
    const win = fromEvent(e);
    if (!win || !trusted(e) || process.platform !== 'darwin') return;
    if (typeof height !== 'number' || !Number.isFinite(height) || height < 20 || height > 120) return;
    // The controls are 14px tall; 28px is the standard bar, where macOS's own spot is right.
    win.setWindowButtonPosition(height === 28 ? null : { x: 16, y: Math.round((height - 14) / 2) });
  });
  ipcMain.on(Channels.setMenu, (e, items: readonly NativeMenuItem[]) => {
    const win = fromEvent(e);
    if (!win || !trusted(e)) return;
    const template = toTemplate(items, win);
    if (process.platform === 'darwin') template.unshift({ role: 'appMenu' });
    Menu.setApplicationMenu(Menu.buildFromTemplate(template));
  });
  ipcMain.on(Channels.closeAck, (e) => {
    const win = fromEvent(e);
    const state = win && windows.get(win.id);
    if (!state || !trusted(e)) return;
    if (state.closeTimer) clearTimeout(state.closeTimer);
    state.closeTimer = null;
  });
  ipcMain.on(Channels.confirmClose, (e) => {
    const win = fromEvent(e);
    const state = win && windows.get(win.id);
    if (!win || !state || !trusted(e)) return;
    state.allowClose = true;
    win.close();
  });
  ipcMain.on(Channels.cancelClose, (e) => {
    if (trusted(e)) quitting = false;
  });
  ipcMain.on(Channels.watch, (e, p: unknown) => {
    const win = fromEvent(e);
    const state = win && windows.get(win.id);
    if (!win || !state || !trusted(e)) return;
    stopWatching(state);
    if (!grants.has(p)) return;
    // Watch the folder, so the watch survives atomic replaces (ours and other apps').
    state.stopWatch = watchFile(p, () => {
      void stat(p)
        .then((s) => {
          const known = knownMtime.get(p);
          if (known !== undefined && Math.abs(s.mtimeMs - known) <= 1) return; // our own save
          if (!win.isDestroyed()) win.webContents.send(Channels.externalChange, p);
        })
        .catch(() => {}); // gone for now (mid-replace or deleted): a later event will tell
    });
  });
}

/** Open files in the first window, or remember them until one has loaded. */
function openInWindow(paths: readonly string[]): void {
  const win = BrowserWindow.getAllWindows()[0];
  if (win && !win.webContents.isLoading()) {
    for (const p of paths) win.webContents.send(Channels.openFile, p);
    return;
  }
  pendingOpen.push(...paths);
  if (!win && app.isReady()) createWindow();
}

app.on('open-file', (e, p) => {
  e.preventDefault();
  openInWindow(grantAll([path.resolve(p)]));
});

app.on('second-instance', (_e, argv, workingDirectory) => {
  const files = grantAll(spreadsheetArgs(argv.slice(1), workingDirectory));
  const win = BrowserWindow.getAllWindows()[0];
  if (win) {
    if (win.isMinimized()) win.restore();
    win.focus();
  }
  if (files.length > 0 || !win) openInWindow(files);
});

app.on('before-quit', () => {
  quitting = true;
});

void app.whenReady().then(async () => {
  if (!single) return; // quitting: the running instance has the files
  protocol.handle('app', (req) => {
    const { pathname } = new URL(req.url);
    const file = path.normalize(path.join(webRoot, decodeURIComponent(pathname)));
    if (!file.startsWith(webRoot + path.sep)) return new Response('Forbidden', { status: 403 });
    return net.fetch(pathToFileURL(file).toString());
  });
  registerIpc();
  await grants.load();
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  // macOS keeps the app running when its last window closes, except after Cmd+Q.
  if (process.platform !== 'darwin' || quitting) app.quit();
});
