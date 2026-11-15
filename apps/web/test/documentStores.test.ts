import {
  BrowserStore,
  ConflictError,
  NasStore,
  NativeStore,
  type CellestraNative,
  type DocumentRef,
} from '@cellestra/ui';
import { describe, expect, it } from 'vitest';
import { DocumentStores, type NasDialogs } from '../src/app/documentStores.ts';

const noDialogs: NasDialogs = {
  choose: () => Promise.resolve(null),
  confirmOverwrite: () => Promise.resolve(false),
};

function stores(options: {
  native?: boolean;
  nasAvailable?: boolean;
  fsAccess?: boolean;
  nas?: NasStore;
  dialogs?: NasDialogs;
}): DocumentStores {
  const win = options.fsAccess ? { showSaveFilePicker: () => Promise.reject(new Error('unused')) } : {};
  return new DocumentStores({
    browser: new BrowserStore(win),
    native: options.native ? new NativeStore({} as CellestraNative) : null,
    nas: options.nas ?? new NasStore('/api', () => Promise.reject(new Error('offline'))),
    nasAvailable: () => options.nasAvailable ?? false,
    nasDialogs: options.dialogs ?? noDialogs,
  });
}

const ref = (store: DocumentRef['store'], extra: Partial<DocumentRef> = {}): DocumentRef => ({
  store,
  name: 'a.ods',
  id: store === 'nas' ? 'docs:a.ods' : 'a.ods',
  version: null,
  ...extra,
});

describe('DocumentStores', () => {
  it('opens and saves new files through Electron first, then the NAS, then this computer', () => {
    expect(stores({ native: true, nasAvailable: true }).primary).toBeInstanceOf(NativeStore);
    const nas = stores({ nasAvailable: true }).primary;
    expect(nas).not.toBeInstanceOf(NativeStore);
    expect(nas).not.toBeInstanceOf(BrowserStore);
    expect(stores({}).primary).toBeInstanceOf(BrowserStore);
  });

  it('finds the store a document lives in, and falls back to the browser for a desktop file in a browser', () => {
    const s = stores({ native: true });
    expect(s.storeFor(ref('native'))).toBe(s.native);
    expect(s.storeFor(ref('nas'))).toBe(s.nas);
    expect(s.storeFor(ref('browser'))).toBe(s.browser);
    expect(stores({}).storeFor(ref('native'))).toBeInstanceOf(BrowserStore);
  });

  it('saves a browser file in place only when it came with a file handle', () => {
    const s = stores({ fsAccess: true });
    expect(s.canSaveInPlace(ref('browser'))).toBe(false);
    expect(s.canSaveInPlace(ref('browser', { handle: {} as FileSystemFileHandle }))).toBe(true);
    expect(stores({}).canSaveInPlace(ref('browser', { handle: {} as FileSystemFileHandle }))).toBe(false);
    expect(s.canSaveInPlace(ref('nas'))).toBe(true);
  });

  it("opens from the NAS at the place the app's folder dialog chose", async () => {
    const nas = new NasStore('/api', () => Promise.reject(new Error('offline')));
    const asked: string[] = [];
    nas.open = (root, path) => {
      asked.push(`${root}:${path}`);
      return Promise.resolve({ ref: ref('nas', { id: `${root}:${path}` }), bytes: new Uint8Array() });
    };
    const s = stores({
      nasAvailable: true,
      nas,
      dialogs: { ...noDialogs, choose: () => Promise.resolve({ root: 'docs', path: 'b.ods' }) },
    });
    const opened = await s.primary.pick();
    expect(asked).toEqual(['docs:b.ods']);
    expect(opened?.ref.id).toBe('docs:b.ods');
  });

  it('asks before Save As overwrites a NAS file, and writes nothing when told no', async () => {
    const nas = new NasStore('/api', () => Promise.reject(new Error('offline')));
    const writes: boolean[] = [];
    nas.saveNew = (root, path, _bytes, overwrite = false) => {
      writes.push(overwrite);
      if (!overwrite) return Promise.reject(new ConflictError('exists', 'There is already a file there.'));
      return Promise.resolve(ref('nas', { id: `${root}:${path}` }));
    };
    let answer = false;
    const s = stores({
      nasAvailable: true,
      nas,
      dialogs: {
        choose: () => Promise.resolve({ root: 'docs', path: 'a.ods' }),
        confirmOverwrite: () => Promise.resolve(answer),
      },
    });
    expect(await s.primary.saveAs('a.ods', new Uint8Array())).toBeNull();
    expect(writes).toEqual([false]);
    answer = true;
    expect((await s.primary.saveAs('a.ods', new Uint8Array()))?.id).toBe('docs:a.ods');
    expect(writes).toEqual([false, false, true]);
  });
});
