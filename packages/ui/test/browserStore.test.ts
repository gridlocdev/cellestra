import { describe, expect, it } from 'vitest';
import { BrowserStore } from '../src/platform/browserStore.ts';
import { ConflictError, type DocumentRef } from '../src/platform/fileStore.ts';

/** A File System Access handle over one in-memory file; `touch` is another app saving it. */
function fakeHandle() {
  let lastModified = 1000;
  let bytes = new Uint8Array();
  const handle = {
    name: 'budget.ods',
    getFile: async () => ({ lastModified, name: 'budget.ods' }),
    createWritable: async () => ({
      write: async (b: Uint8Array) => void (bytes = b),
      close: async () => void (lastModified += 10),
    }),
  };
  return {
    handle: handle as unknown as FileSystemFileHandle,
    touch: () => void (lastModified += 1),
    bytes: () => bytes,
  };
}

const refFor = (handle: FileSystemFileHandle): DocumentRef => ({
  store: 'browser',
  name: 'budget.ods',
  id: 'budget.ods',
  version: '1000',
  handle,
});

describe('BrowserStore.save', () => {
  it('refuses to overwrite a file another app changed, unless forced', async () => {
    const store = new BrowserStore({});
    const file = fakeHandle();
    const ref = refFor(file.handle);
    file.touch();
    await expect(store.save(ref, new Uint8Array([1]))).rejects.toBeInstanceOf(ConflictError);
    expect(file.bytes()).toHaveLength(0);
    await store.save(ref, new Uint8Array([1]), { force: true });
    expect(file.bytes()).toEqual(new Uint8Array([1]));
  });

  it('moves the version forward, so the next save is not a false conflict', async () => {
    const store = new BrowserStore({});
    const file = fakeHandle();
    const saved = await store.save(refFor(file.handle), new Uint8Array([1]));
    expect(saved.version).toBe('1010');
    await expect(store.save(saved, new Uint8Array([2]))).resolves.toBeTruthy();
  });
});
