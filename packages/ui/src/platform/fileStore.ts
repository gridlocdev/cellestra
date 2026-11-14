import type { LockOwner } from './lockFile.ts';

/**
 * Where a document lives. The UI only ever talks to a FileStore; it never knows whether bytes came
 * from the browser, the NAS server, or Electron's native file system.
 */
export type StoreKind = 'browser' | 'nas' | 'native';

/** An opened document's identity and the version it was read at (for conflict-on-save). */
export interface DocumentRef {
  readonly store: StoreKind;
  readonly name: string;
  /** Store-specific locator: a native path, `root:path` on the NAS, or an opaque browser key. */
  readonly id: string;
  /** Etag (NAS) or mtime (native) at open/save; null when the store can't tell. */
  readonly version: string | null;
  /** Browser-only: the File System Access handle, when the browser granted one. */
  readonly handle?: FileSystemFileHandle;
  /**
   * Browser-only: the bytes went out as a download. The browser doesn't say whether the user
   * kept the file, so this is a copy, not a save: the document stays modified.
   */
  readonly downloaded?: boolean;
}

export interface OpenedDocument {
  readonly ref: DocumentRef;
  readonly bytes: Uint8Array;
  /** LibreOffice had the file open when it was read (a lock file beside it), where the store can tell. */
  readonly lock?: LockOwner;
}

export interface StoreCapabilities {
  /** Save writes back to the same file (false → Save downloads a copy). */
  readonly saveInPlace: boolean;
  /** The store can list folders (NAS) — the UI shows its own browser dialog. */
  readonly browse: boolean;
  /** The store can notice the file changing underneath us. */
  readonly watch: boolean;
}

export interface FileStore {
  readonly kind: StoreKind;
  readonly label: string;
  readonly capabilities: StoreCapabilities;
  /** Ask the user for a file (native/browser picker). Null when cancelled. */
  pick(): Promise<OpenedDocument | null>;
  /** Re-read a document (revert, external change). */
  reload(ref: DocumentRef): Promise<OpenedDocument>;
  /** Write back over `ref`. Throws ConflictError when it changed since `ref.version`, unless forced. */
  /**
   * `force` overwrites a file changed since it was opened. `background` (auto-save) never asks the
   * user anything, such as a browser permission prompt; it fails with NotAllowedError instead.
   */
  save(
    ref: DocumentRef,
    bytes: Uint8Array,
    opts?: { force?: boolean; background?: boolean },
  ): Promise<DocumentRef>;
  /** Ask where to save a new copy. Null when cancelled. */
  saveAs(suggestedName: string, bytes: Uint8Array): Promise<DocumentRef | null>;
}

/** The file changed (or is locked) since we opened it. `force` save overwrites anyway. */
export class ConflictError extends Error {
  readonly reason: 'changed' | 'locked' | 'exists';

  constructor(reason: 'changed' | 'locked' | 'exists', message: string) {
    super(message);
    this.name = 'ConflictError';
    this.reason = reason;
  }
}

export const SPREADSHEET_TYPES = {
  'application/vnd.oasis.opendocument.spreadsheet': ['.ods'],
  'text/csv': ['.csv'],
  'text/tab-separated-values': ['.tsv'],
} as const;

export const ACCEPT = '.ods,.csv,.tsv';

export function mimeFor(name: string): string {
  const lower = name.toLowerCase();
  if (lower.endsWith('.csv')) return 'text/csv';
  if (lower.endsWith('.tsv')) return 'text/tab-separated-values';
  return 'application/vnd.oasis.opendocument.spreadsheet';
}
