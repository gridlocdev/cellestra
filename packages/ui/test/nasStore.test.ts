import { describe, expect, it } from 'vitest';
import { latestOnly } from '../src/platform/nasStore.ts';

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
