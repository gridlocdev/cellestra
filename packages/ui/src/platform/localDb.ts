import type { DocumentRef } from './fileStore.ts';

/**
 * Small IndexedDB store for recent files and crash-recovery snapshots. Browser file handles are
 * structured-cloneable, so a recent browser file reopens without a picker (after a permission prompt).
 */
const DB = 'cellestra';
const VERSION = 1;
const RECENT_LIMIT = 10;

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains('recent')) db.createObjectStore('recent', { keyPath: 'key' });
      if (!db.objectStoreNames.contains('recovery')) db.createObjectStore('recovery', { keyPath: 'key' });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error ?? new Error('IndexedDB unavailable'));
  });
}

async function tx<T>(
  store: string,
  mode: IDBTransactionMode,
  fn: (s: IDBObjectStore) => IDBRequest<T>,
): Promise<T> {
  const db = await openDb();
  return new Promise<T>((resolve, reject) => {
    let settled = false;
    const fail = (e: unknown) => {
      if (settled) return;
      settled = true;
      db.close();
      reject(e instanceof Error || e instanceof DOMException ? e : new Error('IndexedDB transaction failed'));
    };
    let t: IDBTransaction;
    let req: IDBRequest<T>;
    try {
      t = db.transaction(store, mode);
      req = fn(t.objectStore(store));
    } catch (e) {
      fail(e);
      return;
    }
    t.oncomplete = () => {
      if (settled) return;
      settled = true;
      db.close();
      resolve(req.result);
    };
    t.onerror = () => fail(t.error ?? req.error);
    t.onabort = () => fail(t.error ?? new DOMException('IndexedDB transaction aborted', 'AbortError'));
  });
}

export interface RecentEntry {
  readonly key: string;
  readonly ref: DocumentRef;
  readonly opened: number;
}

const keyOf = (ref: DocumentRef) => `${ref.store}:${ref.id}`;

export async function recentFiles(): Promise<RecentEntry[]> {
  try {
    const all = await tx<RecentEntry[]>('recent', 'readonly', (s) => s.getAll() as IDBRequest<RecentEntry[]>);
    return all.sort((a, b) => b.opened - a.opened).slice(0, RECENT_LIMIT);
  } catch {
    return [];
  }
}

/**
 * A browser file has only a bare name, so two `data.csv` files from different folders would share
 * a key. Reuse the key of an entry whose handle is the same file, else give it a fresh one.
 */
async function recentKey(ref: DocumentRef, all: readonly RecentEntry[]): Promise<string> {
  if (ref.store !== 'browser' || !ref.handle) return keyOf(ref);
  for (const e of all) {
    if (e.ref.store !== 'browser' || !e.ref.handle || e.ref.name !== ref.name) continue;
    if (await ref.handle.isSameEntry(e.ref.handle).catch(() => false)) return e.key;
  }
  return `${keyOf(ref)}#${randomId()}`;
}

export async function rememberRecent(ref: DocumentRef): Promise<void> {
  // An uploaded browser file (no handle) cannot be reopened, so it is not worth listing.
  if (ref.store === 'browser' && !ref.handle) return;
  try {
    const before = await tx<RecentEntry[]>(
      'recent',
      'readonly',
      (s) => s.getAll() as IDBRequest<RecentEntry[]>,
    );
    const key = await recentKey(ref, before);
    await tx('recent', 'readwrite', (s) => s.put({ key, ref, opened: Date.now() }));
    const all = await tx<RecentEntry[]>('recent', 'readonly', (s) => s.getAll() as IDBRequest<RecentEntry[]>);
    const stale = all.sort((a, b) => b.opened - a.opened).slice(RECENT_LIMIT);
    for (const e of stale) await tx('recent', 'readwrite', (s) => s.delete(e.key));
  } catch {
    // private browsing or storage disabled — recents are a convenience
  }
}

export async function forgetRecent(entry: RecentEntry): Promise<void> {
  await tx('recent', 'readwrite', (s) => s.delete(entry.key)).catch(() => {});
}

export interface RecoverySnapshot {
  /** The tab or window that wrote it (`recoveryKey`); legacy snapshots used `'current'`. */
  readonly key: string;
  readonly name: string;
  readonly ref: DocumentRef | null;
  readonly bytes: Uint8Array;
  readonly saved: number;
}

function randomId(): string {
  return (
    globalThis.crypto?.randomUUID?.() ?? `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`
  );
}

/**
 * This tab's (or window's) recovery slot. Each tab writes and clears only its own snapshot, and
 * holds a Web Lock named after it for its lifetime, so another tab can tell a live tab's snapshot
 * from one left behind by a crash.
 */
export const recoveryKey = `session-${randomId()}`;
const LOCK_PREFIX = 'cellestra-recovery:';
let claimed = false;

interface LockManagerLike {
  request(name: string, cb: () => Promise<unknown>): Promise<unknown>;
  query(): Promise<{ held?: { name?: string }[] }>;
}
function locks(): LockManagerLike | null {
  return (globalThis.navigator as { locks?: LockManagerLike } | undefined)?.locks ?? null;
}

function claimRecoverySlot(): void {
  if (claimed) return;
  claimed = true;
  void locks()
    ?.request(LOCK_PREFIX + recoveryKey, () => new Promise<never>(() => {}))
    .catch(() => {});
}

/**
 * Keep the latest unsaved state so a crash or closed tab loses at most a few seconds: the app
 * snapshots a few seconds after editing pauses, when the tab is hidden, and every 30 s.
 */
export async function saveRecovery(name: string, ref: DocumentRef | null, bytes: Uint8Array): Promise<void> {
  claimRecoverySlot();
  const snap: RecoverySnapshot = { key: recoveryKey, name, ref, bytes, saved: Date.now() };
  await tx('recovery', 'readwrite', (s) => s.put(snap)).catch(() => {});
}

/**
 * Snapshots no live tab owns (left by a crash or a closed tab), newest first. Without Web Locks
 * every other tab's snapshot is offered.
 */
export async function orphanedRecoveries(): Promise<RecoverySnapshot[]> {
  claimRecoverySlot();
  const all = await tx<RecoverySnapshot[]>(
    'recovery',
    'readonly',
    (s) => s.getAll() as IDBRequest<RecoverySnapshot[]>,
  ).catch(() => [] as RecoverySnapshot[]);
  const held = new Set<string>();
  const lm = locks();
  if (lm) {
    const state = await lm.query().catch(() => ({ held: [] }));
    for (const l of state.held ?? [])
      if (l.name?.startsWith(LOCK_PREFIX)) held.add(l.name.slice(LOCK_PREFIX.length));
  }
  return all.filter((s) => s.key !== recoveryKey && !held.has(s.key)).sort((a, b) => b.saved - a.saved);
}

/** Drop a snapshot: this tab's own by default, or an orphan once it was restored or discarded. */
export async function clearRecovery(key: string = recoveryKey): Promise<void> {
  await tx('recovery', 'readwrite', (s) => s.delete(key)).catch(() => {});
}
