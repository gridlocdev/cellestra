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

/**
 * Files in the NAS folders the Cellestra server exposes (`/api`). Conflict detection uses the
 * etag from open: a save sends If-Match, the server answers 412 if someone else saved first.
 * Picking is the UI's own folder browser (`capabilities.browse`), so `pick()` is not used.
 */
export class NasStore implements FileStore {
  readonly kind = 'nas' as const;
  readonly label = 'NAS';
  readonly capabilities = { saveInPlace: true, browse: true, watch: false };
  private readonly base: string;
  private readonly fetchImpl: typeof fetch;

  constructor(base = './api', fetchImpl: typeof fetch = globalThis.fetch.bind(globalThis)) {
    this.base = base.replace(/\/$/, '');
    this.fetchImpl = fetchImpl;
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
      const res = await this.fetchImpl(`${this.base}/info`, { headers: this.headers() });
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

  pick(): Promise<OpenedDocument | null> {
    return Promise.resolve(null);
  }

  async open(root: string, path: string): Promise<OpenedDocument> {
    const res = await this.request(`/file?${new URLSearchParams({ root, path })}`);
    const name = path.split('/').pop() ?? path;
    return {
      ref: { store: 'nas', name, id: `${root}:${path}`, version: res.headers.get('etag') },
      bytes: new Uint8Array(await res.arrayBuffer()),
    };
  }

  reload(ref: DocumentRef): Promise<OpenedDocument> {
    const { root, path } = splitId(ref.id);
    return this.open(root, path);
  }

  async save(ref: DocumentRef, bytes: Uint8Array, opts: { force?: boolean } = {}): Promise<DocumentRef> {
    const { root, path } = splitId(ref.id);
    return this.put(root, path, bytes, ref.version ? { 'If-Match': ref.version } : {}, opts.force ?? false);
  }

  /** Save a new file into a NAS folder; refuses to overwrite unless `overwrite`. */
  saveNew(root: string, path: string, bytes: Uint8Array, overwrite = false): Promise<DocumentRef> {
    return this.put(root, path, bytes, overwrite ? {} : { 'If-None-Match': '*' }, false);
  }

  /** The UI drives Save As through its own folder browser, then calls `saveNew`. */
  saveAs(): Promise<DocumentRef | null> {
    return Promise.resolve(null);
  }

  private async put(
    root: string,
    path: string,
    bytes: Uint8Array,
    extra: Record<string, string>,
    force: boolean,
  ): Promise<DocumentRef> {
    const q = new URLSearchParams({ root, path, ...(force ? { force: '1' } : {}) });
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
    throw new Error(message);
  }

  private headers(): Record<string, string> {
    const t = this.token;
    return t ? { Authorization: `Bearer ${t}` } : {};
  }
}

function splitId(id: string): { root: string; path: string } {
  const i = id.indexOf(':');
  return { root: id.slice(0, i), path: id.slice(i + 1) };
}
