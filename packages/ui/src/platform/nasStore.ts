import { parseLockFile } from './lockFile.ts';
import { ConflictError, type DocumentRef, type FileStore, type OpenedDocument } from './fileStore.ts';

export interface NasRoot {
  readonly id: string;
  readonly name: string;
}
export interface NasEntry {
  readonly name: string;
  readonly path: string;
  readonly kind: 'file' | 'dir';
  readonly size: number;
  readonly modified: number;
}
export interface NasInfo {
  readonly name: string;
  readonly auth: boolean;
}

/** Needs a token, or the token is wrong. */
export class AuthError extends Error {}

const TOKEN_KEY = 'cellestra.nasToken';
/** The server sends LibreOffice's lock file beside an opened file in this header, percent-encoded. */
const LOCK_HEADER = 'X-Cellestra-Lock';
/** A server that dropped off fails the request instead of leaving a save or an open waiting forever. */
const REQUEST_TIMEOUT_MS = 20_000;
/** The check at start for a server only decides whether NAS is offered, so it gives up sooner. */
const INFO_TIMEOUT_MS = 5000;

/**
 * Files in the NAS folders the Cellestra server exposes (`/api`). Conflict detection uses the
 * etag from open: a save sends If-Match, the server answers 412 if someone else saved first.
 * It has no picker of its own: the app shows its folder browser (`capabilities.browse`), then calls
 * `open` or `saveNew`.
 */
export class NasStore implements FileStore {
  readonly kind = 'nas' as const;
  readonly label = 'NAS';
  readonly capabilities = { saveInPlace: true, browse: true, watch: false };
  private readonly base: string;
  private readonly fetchImpl: typeof fetch;
  /** How long a request may take, body included, before it fails. */
  readonly timeoutMs: number;

  constructor(
    base = './api',
    fetchImpl: typeof fetch = globalThis.fetch.bind(globalThis),
    timeoutMs = REQUEST_TIMEOUT_MS,
  ) {
    this.base = base.replace(/\/$/, '');
    this.fetchImpl = fetchImpl;
    this.timeoutMs = timeoutMs;
  }

  get token(): string | null {
    try {
      return localStorage.getItem(TOKEN_KEY);
    } catch {
      return null;
    }
  }

  set token(value: string | null) {
    try {
      if (value) localStorage.setItem(TOKEN_KEY, value);
      else localStorage.removeItem(TOKEN_KEY);
    } catch {
      // storage disabled: the token lasts only for this request
    }
  }

  /** Null when no server answers (a plain static deployment, or file://). */
  async info(): Promise<NasInfo | null> {
    try {
      const res = await this.fetchImpl(`${this.base}/info`, {
        headers: this.headers(),
        signal: AbortSignal.timeout(Math.min(INFO_TIMEOUT_MS, this.timeoutMs)),
      });
      if (!res.ok || !res.headers.get('content-type')?.includes('json')) return null;
      return (await res.json()) as NasInfo;
    } catch {
      return null;
    }
  }

  roots(): Promise<NasRoot[]> {
    return this.json('/roots');
  }

  list(root: string, path = ''): Promise<NasEntry[]> {
    return this.json(`/list?${new URLSearchParams({ root, path })}`);
  }

  canSaveInPlace(): boolean {
    return true;
  }

  async open(root: string, path: string): Promise<OpenedDocument> {
    const res = await this.request(`/file?${new URLSearchParams({ root, path })}`);
    const name = path.split('/').pop() ?? path;
    const doc = {
      ref: { store: 'nas' as const, name, id: `${root}:${path}`, version: res.headers.get('etag') },
      bytes: new Uint8Array(await res.arrayBuffer()),
    };
    const lock = lockHeader(res.headers.get(LOCK_HEADER));
    return lock === null ? doc : { ...doc, lock: parseLockFile(lock) };
  }

  /**
   * The file's ETag now, from a HEAD request (no download), so an open document can notice someone
   * else saving it. Null when the file is gone.
   */
  async currentVersion(id: string): Promise<string | null> {
    const { root, path } = splitId(id);
    try {
      const res = await this.request(`/file?${new URLSearchParams({ root, path })}`, {
        method: 'HEAD',
        cache: 'no-store',
      });
      return res.headers.get('etag');
    } catch (e) {
      if (e instanceof Error && e.name === 'NotFoundError') return null;
      throw e;
    }
  }

  reload(ref: DocumentRef): Promise<OpenedDocument> {
    const { root, path } = splitId(ref.id);
    return this.open(root, path);
  }

  async save(
    ref: DocumentRef,
    bytes: Uint8Array,
    opts: { force?: boolean; ignoreLock?: boolean } = {},
  ): Promise<DocumentRef> {
    const { root, path } = splitId(ref.id);
    const headers: Record<string, string> = ref.version ? { 'If-Match': ref.version } : {};
    return this.put(root, path, bytes, headers, opts.force ?? false, opts.ignoreLock ?? false);
  }

  /** Save a new file into a NAS folder; refuses to overwrite unless `overwrite`. */
  saveNew(root: string, path: string, bytes: Uint8Array, overwrite = false): Promise<DocumentRef> {
    return this.put(root, path, bytes, overwrite ? {} : { 'If-None-Match': '*' }, false, false);
  }

  private async put(
    root: string,
    path: string,
    bytes: Uint8Array,
    extra: Record<string, string>,
    force: boolean,
    ignoreLock: boolean,
  ): Promise<DocumentRef> {
    const q = new URLSearchParams({
      root,
      path,
      ...(force ? { force: '1' } : {}),
      ...(ignoreLock ? { ignoreLock: '1' } : {}),
    });
    const res = await this.request(`/file?${q}`, {
      method: 'PUT',
      body: bytes as Uint8Array<ArrayBuffer>,
      headers: { 'Content-Type': 'application/octet-stream', ...extra },
    });
    const { etag } = (await res.json()) as { etag: string };
    return { store: 'nas', name: path.split('/').pop() ?? path, id: `${root}:${path}`, version: etag };
  }

  private async json<T>(path: string): Promise<T> {
    return (await (await this.request(path)).json()) as T;
  }

  private async request(path: string, init: RequestInit = {}): Promise<Response> {
    const res = await this.fetchImpl(this.base + path, {
      ...init,
      headers: { ...this.headers(), ...(init.headers as Record<string, string> | undefined) },
      signal: init.signal ?? AbortSignal.timeout(this.timeoutMs),
    }).catch((e: unknown) => {
      if (e instanceof DOMException && e.name === 'TimeoutError') {
        throw new Error(`The server didn't answer within ${Math.round(this.timeoutMs / 1000)} seconds.`);
      }
      throw e;
    });
    if (res.ok) return res;
    const message = await res
      .json()
      .then((b: { error?: string }) => b.error ?? res.statusText)
      .catch(() => res.statusText);
    if (res.status === 401) throw new AuthError(message);
    if (res.status === 412)
      throw new ConflictError(
        init.method === 'PUT' && message.includes('exists') ? 'exists' : 'changed',
        message,
      );
    if (res.status === 423) throw new ConflictError('locked', message);
    const err = new Error(message);
    if (res.status === 404) err.name = 'NotFoundError';
    throw err;
  }

  private headers(): Record<string, string> {
    const t = this.token;
    return t ? { Authorization: `Bearer ${t}` } : {};
  }
}

/**
 * For requests that can overlap, like folder listings as the user clicks around: each call
 * replaces the one before, and only the newest request's result (or error) is delivered. A slow
 * earlier listing that answers last is dropped instead of overwriting the newer folder.
 */
export function latestOnly(): <T>(
  promise: Promise<T>,
  onValue: (value: T) => void,
  onError: (error: unknown) => void,
) => void {
  let newest = 0;
  return (promise, onValue, onError) => {
    const ticket = ++newest;
    promise.then(
      (value) => {
        if (ticket === newest) onValue(value);
      },
      (error: unknown) => {
        if (ticket === newest) onError(error);
      },
    );
  };
}

/** The lock file text a header carries: null when there is none, '' when it is garbled. */
function lockHeader(value: string | null): string | null {
  if (value === null) return null;
  try {
    return decodeURIComponent(value);
  } catch {
    return '';
  }
}

function splitId(id: string): { root: string; path: string } {
  const i = id.indexOf(':');
  return { root: id.slice(0, i), path: id.slice(i + 1) };
}
