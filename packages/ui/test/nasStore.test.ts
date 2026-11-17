import { describe, expect, it } from 'vitest';
import { NasStore, latestOnly } from '../src/platform/nasStore.ts';

/** A promise and the functions that settle it, so a test picks the order replies arrive in. */
function deferred<T>() {
  let resolve!: (v: T) => void;
  let reject!: (e: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

const flush = () => new Promise((r) => setTimeout(r, 0));

describe('latestOnly', () => {
  it('drops a slow earlier listing that arrives after a newer one', async () => {
    const latest = latestOnly();
    const shown: string[] = [];
    const errors: unknown[] = [];
    const first = deferred<string>();
    const second = deferred<string>();
    latest(
      first.promise,
      (v) => shown.push(v),
      (e) => errors.push(e),
    );
    latest(
      second.promise,
      (v) => shown.push(v),
      (e) => errors.push(e),
    );
    second.resolve('subfolder');
    await flush();
    first.resolve('parent folder');
    await flush();
    expect(shown).toEqual(['subfolder']);
    expect(errors).toEqual([]);
  });

  it('drops a stale failure too, and still delivers the newest one', async () => {
    const latest = latestOnly();
    const errors: unknown[] = [];
    const first = deferred<string>();
    const second = deferred<string>();
    latest(
      first.promise,
      () => {},
      (e) => errors.push(e),
    );
    latest(
      second.promise,
      () => {},
      (e) => errors.push(e),
    );
    first.reject(new Error('old'));
    second.reject(new Error('new'));
    await flush();
    expect(errors.map((e) => (e as Error).message)).toEqual(['new']);
  });
});

describe('NasStore requests', () => {
  it('fails a request the server never answers, instead of waiting forever', async () => {
    // A server that dropped off: the request hangs until its signal gives up.
    const hang: typeof fetch = (_url, init) =>
      new Promise((_resolve, reject) => {
        init?.signal?.addEventListener('abort', () => reject(init.signal?.reason as Error));
      });
    const store = new NasStore('/api', hang, 20);
    await expect(store.open('docs', 'a.ods')).rejects.toThrow("The server didn't answer within 0 seconds.");
    const ref = { store: 'nas', name: 'a.ods', id: 'docs:a.ods', version: '"e1"' } as const;
    await expect(store.save(ref, new Uint8Array([1]))).rejects.toThrow("didn't answer");
    await expect(store.info()).resolves.toBeNull();
  });
});

describe('NasStore.save', () => {
  it('asks the server to skip the lock only when told to, separately from the version check', async () => {
    const urls: string[] = [];
    const store = new NasStore('/api', (url) => {
      urls.push(String(url));
      return Promise.resolve(new Response(JSON.stringify({ etag: '"e2"' })));
    });
    const ref = { store: 'nas', name: 'a.ods', id: 'docs:a.ods', version: '"e1"' } as const;
    await store.save(ref, new Uint8Array([1]), { force: true });
    await store.save(ref, new Uint8Array([1]), { force: true, ignoreLock: true });
    const query = urls.map((u) => new URL(u, 'http://nas').searchParams);
    expect(query[0]!.get('force')).toBe('1');
    expect(query[0]!.has('ignoreLock')).toBe(false);
    expect(query[1]!.get('ignoreLock')).toBe('1');
  });
});
