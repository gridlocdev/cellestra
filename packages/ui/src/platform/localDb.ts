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
    const t = db.transaction(store, mode);
    const req = fn(t.objectStore(store));
    t.oncomplete = () => {
      db.close();
      resolve(req.result);
    };
    t.onerror = () => {
      db.close();
      reject(t.error ?? new Error('IndexedDB transaction failed'));
    };
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

export async function rememberRecent(ref: DocumentRef): Promise<void> {
  try {
    await tx('recent', 'readwrite', (s) => s.put({ key: keyOf(ref), ref, opened: Date.now() }));
    const all = await tx<RecentEntry[]>('recent', 'readonly', (s) => s.getAll() as IDBRequest<RecentEntry[]>);
    const stale = all.sort((a, b) => b.opened - a.opened).slice(RECENT_LIMIT);
    for (const e of stale) await tx('recent', 'readwrite', (s) => s.delete(e.key));
  } catch {
    // private browsing or storage disabled — recents are a convenience
  }
}

export async function forgetRecent(ref: DocumentRef): Promise<void> {
  await tx('recent', 'readwrite', (s) => s.delete(keyOf(ref))).catch(() => {});
}

export interface RecoverySnapshot {
  readonly key: 'current';
  readonly name: string;
  readonly ref: DocumentRef | null;
  readonly bytes: Uint8Array;
  readonly saved: number;
}

/** Keep the latest unsaved state so a crash or closed tab loses at most a few seconds. */
export async function saveRecovery(name: string, ref: DocumentRef | null, bytes: Uint8Array): Promise<void> {
  const snap: RecoverySnapshot = { key: 'current', name, ref, bytes, saved: Date.now() };
  await tx('recovery', 'readwrite', (s) => s.put(snap)).catch(() => {});
}

export async function loadRecovery(): Promise<RecoverySnapshot | null> {
  return tx<RecoverySnapshot | undefined>(
    'recovery',
    'readonly',
    (s) => s.get('current') as IDBRequest<RecoverySnapshot | undefined>,
  )
    .then((r) => r ?? null)
    .catch(() => null);
}

export async function clearRecovery(): Promise<void> {
  await tx('recovery', 'readwrite', (s) => s.delete('current')).catch(() => {});
}
