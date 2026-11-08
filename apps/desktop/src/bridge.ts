/**
 * The contract between the Electron main process and the renderer, shared by `main.ts` and
 * `preload.cts`. The renderer sees `window.cellestraNative` (typed again on the UI side in
 * `packages/ui/src/platform/native.ts` — keep the two in step).
 */

export interface NativeFile {
  readonly path: string;
  readonly name: string;
  readonly bytes: Uint8Array;
  /** mtime in ms when read; passed back on save to detect outside edits. */
  readonly modified: number;
}

export interface SavedFile {
  readonly path: string;
  readonly name: string;
  readonly modified: number;
}

/** A menu the renderer builds from its command catalog; `command` ids come back via onCommand. */
export interface NativeMenuItem {
  readonly label?: string;
  readonly command?: string;
  readonly accelerator?: string;
  readonly enabled?: boolean;
  readonly checked?: boolean;
  readonly role?: string;
  readonly type?: 'separator' | 'checkbox' | 'normal';
  readonly submenu?: readonly NativeMenuItem[];
}

export interface DocumentState {
  readonly title: string;
  readonly dirty: boolean;
  readonly path: string | null;
}

export const Channels = {
  openDialog: 'cellestra:open-dialog',
  readFile: 'cellestra:read-file',
  saveFile: 'cellestra:save-file',
  saveDialog: 'cellestra:save-dialog',
  setDocumentState: 'cellestra:document-state',
  setMenu: 'cellestra:set-menu',
  setTitleBarHeight: 'cellestra:title-bar-height',
  confirmClose: 'cellestra:confirm-close',
  /** The renderer kept the window open (the user cancelled the unsaved-changes prompt). */
  cancelClose: 'cellestra:cancel-close',
  /** Sent by the preload as soon as a close request arrives: proves the renderer is alive. */
  closeAck: 'cellestra:close-ack',
  watch: 'cellestra:watch',
  // main → renderer
  command: 'cellestra:command',
  openFile: 'cellestra:open-file',
  closeRequested: 'cellestra:close-requested',
  externalChange: 'cellestra:external-change',
} as const;

/** Thrown (as an IPC error message prefix) when the file changed on disk since it was opened. */
export const CONFLICT = 'CELLESTRA_CONFLICT';
