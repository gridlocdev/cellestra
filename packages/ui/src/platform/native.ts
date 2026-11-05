// Mirrors apps/desktop/src/bridge.ts — the preload exposes this as `window.cellestraNative`.
export interface NativeFile {
  readonly path: string;
  readonly name: string;
  readonly bytes: Uint8Array;
  readonly modified: number;
}
export interface SavedFile {
  readonly path: string;
  readonly name: string;
  readonly modified: number;
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
export interface CellestraNative {
  readonly platform: string;
  openDialog(): Promise<NativeFile | null>;
  readFile(path: string): Promise<NativeFile>;
  saveFile(
    path: string,
    bytes: Uint8Array,
    expectedModified: number | null,
    force: boolean,
  ): Promise<SavedFile>;
  saveDialog(suggestedName: string, bytes: Uint8Array): Promise<SavedFile | null>;
  setDocumentState(state: DocumentState): void;
  setMenu(menu: readonly NativeMenuItem[]): void;
  confirmClose(): void;
  /** The close (or quit) was cancelled; the window stays open. */
  cancelClose(): void;
  watch(path: string | null): void;
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
