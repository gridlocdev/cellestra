// Mirrors apps/desktop/src/bridge.ts — the preload exposes this as `window.cellestraNative`.
export interface NativeFile {
  readonly path: string;
  readonly name: string;
  readonly bytes: Uint8Array;
  readonly modified: number;
  /** A hash of the contents, passed back on save to detect outside edits. */
  readonly version: string;
  /** LibreOffice's lock file beside the file, as text. Optional: older desktop builds omit it. */
  readonly lock?: string | null;
}
export interface SavedFile {
  readonly path: string;
  readonly name: string;
  readonly modified: number;
  readonly version: string;
}
export interface NativeMenuItem {
  readonly label?: string;
  readonly command?: string;
  readonly accelerator?: string;
  readonly enabled?: boolean;
  readonly checked?: boolean;
  readonly role?: string;
  readonly type?: 'normal' | 'separator' | 'checkbox' | 'submenu';
  readonly submenu?: readonly NativeMenuItem[];
}
export interface DocumentState {
  readonly title: string;
  readonly dirty: boolean;
  readonly path: string | null;
}
export interface PrintPaper {
  readonly width: number;
  readonly height: number;
}
/** A key the desktop app held back during a lookup, with the fields shortcuts read. */
export interface CapturedKey {
  readonly key: string;
  readonly code: string;
  readonly shiftKey: boolean;
  readonly ctrlKey: boolean;
  readonly altKey: boolean;
  readonly metaKey: boolean;
}
export interface CellestraNative {
  readonly platform: string;
  openDialog(): Promise<NativeFile | null>;
  readFile(path: string): Promise<NativeFile>;
  saveFile(
    path: string,
    bytes: Uint8Array,
    expectedVersion: string | null,
    force: boolean,
  ): Promise<SavedFile>;
  saveDialog(suggestedName: string, bytes: Uint8Array): Promise<SavedFile | null>;
  setDocumentState(state: DocumentState): void;
  setMenu(menu: readonly NativeMenuItem[]): void;
  /**
   * macOS: the height of the page's own title bar, so the window controls sit centred in it.
   * Optional so an older desktop build without it still works.
   */
  setTitleBarHeight?(height: number): void;
  confirmClose(): void;
  /** The close (or quit) was cancelled; the window stays open. */
  cancelClose(): void;
  watch(path: string | null): void;
  /** Print a document through the system print dialog; false when cancelled. Optional: older builds. */
  print?(html: string, paper: PrintPaper): Promise<boolean>;
  /** Save a document as PDF where the user picks; the path, or null when cancelled. */
  exportPdf?(html: string, suggestedName: string): Promise<string | null>;
  /**
   * While on, the window holds keys back from the menu bar and the page and sends them to
   * `onCapturedKey` instead (the Keyboard Shortcuts lookup). Escape always gets through and turns
   * it off. Optional: older desktop builds.
   */
  captureKeys?(on: boolean): void;
  onCapturedKey?(cb: (key: CapturedKey) => void): () => void;
  onCommand(cb: (id: string) => void): () => void;
  onOpenFile(cb: (path: string) => void): () => void;
  onCloseRequested(cb: () => void): () => void;
  onExternalChange(cb: (path: string) => void): () => void;
}

export const CONFLICT = 'CELLESTRA_CONFLICT';

/** The Electron bridge, or null in a plain browser. */
export function nativeBridge(): CellestraNative | null {
  return (globalThis as { cellestraNative?: CellestraNative }).cellestraNative ?? null;
}
