import {
  ConflictError,
  type BrowserStore,
  type DocumentRef,
  type FilePicker,
  type FileStore,
  type NasStore,
  type NativeStore,
  type OpenedDocument,
} from '@cellestra/ui';

/** A place in a NAS folder the user chose in the app's folder dialog. */
export interface NasLocation {
  readonly root: string;
  readonly path: string;
}

/** The questions the NAS picker asks through the app's own dialogs. */
export interface NasDialogs {
  /** Show the folder browser. Null when cancelled. */
  choose(mode: 'open' | 'save', suggestedName?: string): Promise<NasLocation | null>;
  /** A file of that name is already there: whether to overwrite it. */
  confirmOverwrite(e: ConflictError): Promise<boolean>;
}

export interface DocumentStoresOptions {
  readonly browser: BrowserStore;
  /** Electron's file system; null in a browser. */
  readonly native: NativeStore | null;
  readonly nas: NasStore;
  /** Whether the NAS server answered, so Open and Save As go there. Read at each call. */
  readonly nasAvailable: () => boolean;
  readonly nasDialogs: NasDialogs;
}

/**
 * The file stores a document can live in, and which one Open and Save As use. Callers ask for a
 * store by document or for the primary picker; they never switch on the store's kind.
 */
export class DocumentStores {
  readonly browser: BrowserStore;
  readonly native: NativeStore | null;
  nas: NasStore;
  readonly #nasAvailable: () => boolean;
  readonly #nasPicker: FilePicker;

  constructor(options: DocumentStoresOptions) {
    this.browser = options.browser;
    this.native = options.native;
    this.nas = options.nas;
    this.#nasAvailable = options.nasAvailable;
    this.#nasPicker = nasPicker(() => this.nas, options.nasDialogs);
  }

  /** The store that holds `ref`. A native file seen in a browser falls back to the browser store. */
  storeFor(ref: DocumentRef): FileStore {
    if (ref.store === 'native' && this.native) return this.native;
    if (ref.store === 'nas') return this.nas;
    return this.browser;
  }

  /** Where Open and Save As go: Electron's file system, else the NAS, else this computer. */
  get primary(): FilePicker {
    if (this.native) return this.native;
    return this.#nasAvailable() ? this.#nasPicker : this.browser;
  }

  /** Open a document again from where it was last seen (a recent file). */
  reopen(ref: DocumentRef): Promise<OpenedDocument> {
    return this.storeFor(ref).reload(ref);
  }

  /** Whether Save writes `ref` back in place; when not, Save asks where (Save As). */
  canSaveInPlace(ref: DocumentRef): boolean {
    const store = this.storeFor(ref);
    return store.capabilities.saveInPlace && store.canSaveInPlace(ref);
  }
}

/** Open and Save As on the NAS: the app's folder dialog picks the place, the store reads or writes it. */
export function nasPicker(nas: () => NasStore, dialogs: NasDialogs): FilePicker {
  return {
    async pick() {
      const where = await dialogs.choose('open');
      return where ? nas().open(where.root, where.path) : null;
    },
    async saveAs(name, bytes) {
      const where = await dialogs.choose('save', name);
      if (!where) return null;
      try {
        return await nas().saveNew(where.root, where.path, bytes);
      } catch (e) {
        if (!(e instanceof ConflictError) || e.reason !== 'exists') throw e;
        return (await dialogs.confirmOverwrite(e))
          ? nas().saveNew(where.root, where.path, bytes, true)
          : null;
      }
    },
  };
}
