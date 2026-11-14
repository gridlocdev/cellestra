import { CONFLICT, type CellestraNative, type NativeFile, type SavedFile } from './native.ts';
import { parseLockFile } from './lockFile.ts';
import { ConflictError, type DocumentRef, type FileStore, type OpenedDocument } from './fileStore.ts';

/** Files on disk through Electron's main process. Version = a hash of the contents, checked again before each save. */
export class NativeStore implements FileStore {
  readonly kind = 'native' as const;
  readonly label = 'Files';
  readonly capabilities = { saveInPlace: true, browse: false, watch: true };
  readonly bridge: CellestraNative;

  constructor(bridge: CellestraNative) {
    this.bridge = bridge;
  }

  async pick(): Promise<OpenedDocument | null> {
    const f = await this.bridge.openDialog();
    return f ? opened(f) : null;
  }

  async openPath(path: string): Promise<OpenedDocument> {
    return opened(await this.bridge.readFile(path));
  }

  reload(ref: DocumentRef): Promise<OpenedDocument> {
    return this.openPath(ref.id);
  }

  async save(ref: DocumentRef, bytes: Uint8Array, opts: { force?: boolean } = {}): Promise<DocumentRef> {
    try {
      return toRef(await this.bridge.saveFile(ref.id, bytes, ref.version, opts.force ?? false));
    } catch (e) {
      // Errors cross IPC as plain messages; the main process prefixes conflicts with CONFLICT.
      if (e instanceof Error && e.message.includes(CONFLICT))
        throw new ConflictError('changed', 'The file changed on disk since you opened it.');
      throw e;
    }
  }

  async saveAs(suggestedName: string, bytes: Uint8Array): Promise<DocumentRef | null> {
    const saved = await this.bridge.saveDialog(suggestedName, bytes);
    return saved ? toRef(saved) : null;
  }
}

function opened(f: NativeFile): OpenedDocument {
  const doc = { ref: toRef(f), bytes: f.bytes };
  return f.lock ? { ...doc, lock: parseLockFile(f.lock) } : doc;
}

function toRef(f: SavedFile | NativeFile): DocumentRef {
  return { store: 'native', name: f.name, id: f.path, version: f.version };
}
