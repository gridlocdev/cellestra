// Sandboxed preload: may only require('electron'), so channel names are inlined here.
// Keep them in step with `Channels` in bridge.ts.
import type { DocumentState, NativeFile, NativeMenuItem, SavedFile } from './bridge.ts';

// A sandboxed preload is CommonJS and can only reach Electron through require().
// eslint-disable-next-line @typescript-eslint/no-require-imports, @typescript-eslint/consistent-type-imports
const { contextBridge, ipcRenderer } = require('electron') as typeof import('electron');

function on(channel: string, cb: (...args: never[]) => void): () => void {
  const listener = (_e: unknown, ...args: unknown[]) => (cb as (...a: unknown[]) => void)(...args);
  ipcRenderer.on(channel, listener);
  return () => ipcRenderer.removeListener(channel, listener);
}

contextBridge.exposeInMainWorld('cellestraNative', {
  platform: process.platform,
  openDialog: (): Promise<NativeFile | null> => ipcRenderer.invoke('cellestra:open-dialog'),
  readFile: (path: string): Promise<NativeFile> => ipcRenderer.invoke('cellestra:read-file', path),
  saveFile: (
    path: string,
    bytes: Uint8Array,
    expectedModified: number | null,
    force: boolean,
  ): Promise<SavedFile> => ipcRenderer.invoke('cellestra:save-file', path, bytes, expectedModified, force),
  saveDialog: (suggestedName: string, bytes: Uint8Array): Promise<SavedFile | null> =>
    ipcRenderer.invoke('cellestra:save-dialog', suggestedName, bytes),
  setDocumentState: (state: DocumentState): void => ipcRenderer.send('cellestra:document-state', state),
  setMenu: (menu: readonly NativeMenuItem[]): void => ipcRenderer.send('cellestra:set-menu', menu),
  confirmClose: (): void => ipcRenderer.send('cellestra:confirm-close'),
  cancelClose: (): void => ipcRenderer.send('cellestra:cancel-close'),
  watch: (path: string | null): void => ipcRenderer.send('cellestra:watch', path),
  onCommand: (cb: (id: string) => void) => on('cellestra:command', cb),
  onOpenFile: (cb: (path: string) => void) => on('cellestra:open-file', cb),
  onCloseRequested: (cb: () => void) =>
    on('cellestra:close-requested', () => {
      // Tell the main process the page is alive before it starts asking about unsaved changes.
      ipcRenderer.send('cellestra:close-ack');
      cb();
    }),
  onExternalChange: (cb: (path: string) => void) => on('cellestra:external-change', cb),
});
