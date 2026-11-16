import { after } from '../schedule.ts';
import {
  ACCEPT,
  ConflictError,
  mimeFor,
  SPREADSHEET_TYPES,
  type DocumentRef,
  type FilePicker,
  type FileStore,
  type OpenedDocument,
} from './fileStore.ts';

// File System Access API (Chromium). Typed locally — lib.dom only has the handle types.
interface PickerWindow {
  showOpenFilePicker?: (opts: object) => Promise<FileSystemFileHandle[]>;
  showSaveFilePicker?: (opts: object) => Promise<FileSystemFileHandle>;
}
interface WritableHandle extends FileSystemFileHandle {
  queryPermission?(d: { mode: 'read' | 'readwrite' }): Promise<PermissionState>;
  requestPermission?(d: { mode: 'read' | 'readwrite' }): Promise<PermissionState>;
}

/**
 * Make sure the page may use a handle (one restored from IndexedDB starts at 'prompt'). Asking
 * needs a user gesture, so `ask` is false for background work: then it only checks.
 */
async function ensurePermission(
  handle: FileSystemFileHandle,
  mode: 'read' | 'readwrite',
  ask = true,
): Promise<void> {
  const h = handle as WritableHandle;
  if (!h.queryPermission) return;
  if ((await h.queryPermission({ mode })) === 'granted') return;
  if (ask && (await h.requestPermission?.({ mode })) === 'granted') return;
  throw new DOMException(
    `Cellestra was not allowed to ${mode === 'read' ? 'open' : 'save'} ${handle.name}.`,
    'NotAllowedError',
  );
}

const pickerTypes = [{ description: 'Spreadsheets', accept: SPREADSHEET_TYPES }];

function isAbort(e: unknown): boolean {
  return e instanceof DOMException && e.name === 'AbortError';
}

/** Files on this computer, through the browser: File System Access where available, else upload/download. */
export class BrowserStore implements FileStore, FilePicker {
  readonly kind = 'browser' as const;
  readonly label = 'This computer';
  readonly capabilities;
  private readonly win: PickerWindow;

  constructor(win: PickerWindow = globalThis as PickerWindow) {
    this.win = win;
    const fsAccess = typeof win.showSaveFilePicker === 'function';
    this.capabilities = { saveInPlace: fsAccess, browse: false, watch: false };
  }

  /** Only a file opened through a File System Access handle can be written back. */
  canSaveInPlace(ref: DocumentRef): boolean {
    return this.capabilities.saveInPlace && ref.handle !== undefined;
  }

  async pick(): Promise<OpenedDocument | null> {
    if (this.win.showOpenFilePicker) {
      try {
        const [handle] = await this.win.showOpenFilePicker({ types: pickerTypes, multiple: false });
        return handle ? this.fromHandle(handle) : null;
      } catch (e) {
        if (isAbort(e)) return null;
        throw e;
      }
    }
    const file = await chooseFile();
    return file ? fromFile(file) : null;
  }

  /** Open a File the user dropped or chose some other way. */
  async openFile(file: File, handle?: FileSystemFileHandle): Promise<OpenedDocument> {
    return handle ? this.fromHandle(handle) : fromFile(file);
  }

  async reload(ref: DocumentRef): Promise<OpenedDocument> {
    if (!ref.handle) throw new Error('This file was uploaded; open it again to reload it.');
    return this.fromHandle(ref.handle);
  }

  async save(
    ref: DocumentRef,
    bytes: Uint8Array,
    opts?: { force?: boolean; background?: boolean },
  ): Promise<DocumentRef> {
    if (ref.handle) {
      await ensurePermission(ref.handle, 'readwrite', !opts?.background);
      // Another app (or tab) saved the file since we read it: don't write over its changes unasked.
      if (!opts?.force && ref.version !== null) {
        const current = await ref.handle.getFile();
        if (String(current.lastModified) !== ref.version)
          throw new ConflictError('changed', `${ref.name} was changed by another app since it was opened.`);
      }
      const w = await ref.handle.createWritable();
      await w.write(bytes as Uint8Array<ArrayBuffer>);
      await w.close();
      const written = await ref.handle.getFile();
      return { ...ref, version: String(written.lastModified) };
    }
    // No handle (an uploaded or dropped file): the controller sends these through Save As.
    download(ref.name, bytes);
    return { ...ref, downloaded: true };
  }

  async saveAs(suggestedName: string, bytes: Uint8Array): Promise<DocumentRef | null> {
    if (this.win.showSaveFilePicker) {
      try {
        const handle = await this.win.showSaveFilePicker({ suggestedName, types: pickerTypes });
        const ref: DocumentRef = {
          store: 'browser',
          name: handle.name,
          id: handle.name,
          version: null,
          handle,
        };
        return this.save(ref, bytes);
      } catch (e) {
        if (isAbort(e)) return null;
        throw e;
      }
    }
    download(suggestedName, bytes);
    return { store: 'browser', name: suggestedName, id: suggestedName, version: null, downloaded: true };
  }

  private async fromHandle(handle: FileSystemFileHandle): Promise<OpenedDocument> {
    await ensurePermission(handle, 'read');
    const file = await handle.getFile();
    return {
      ref: { store: 'browser', name: file.name, id: file.name, version: String(file.lastModified), handle },
      bytes: new Uint8Array(await file.arrayBuffer()),
    };
  }
}

async function fromFile(file: File): Promise<OpenedDocument> {
  return {
    ref: { store: 'browser', name: file.name, id: file.name, version: String(file.lastModified) },
    bytes: new Uint8Array(await file.arrayBuffer()),
  };
}

function chooseFile(): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = ACCEPT;
    input.addEventListener('change', () => resolve(input.files?.[0] ?? null), { once: true });
    input.addEventListener('cancel', () => resolve(null), { once: true });
    input.click();
  });
}

export function download(name: string, bytes: Uint8Array): void {
  const url = URL.createObjectURL(new Blob([bytes as Uint8Array<ArrayBuffer>], { type: mimeFor(name) }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  after(10_000, () => URL.revokeObjectURL(url));
}
