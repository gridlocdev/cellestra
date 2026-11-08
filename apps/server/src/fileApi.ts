import { createHash, randomBytes } from 'node:crypto';
import { constants as fsConstants } from 'node:fs';
import { chown, mkdir, open, readdir, readFile, realpath, rename, stat, unlink } from 'node:fs/promises';
import path from 'node:path';

/** A named folder on the NAS that the web client may browse and save into. */
export interface Root {
  readonly id: string;
  readonly name: string;
  readonly dir: string;
}

export interface Entry {
  readonly name: string;
  /** Root-relative POSIX path. */
  readonly path: string;
  readonly kind: 'file' | 'dir';
  readonly size: number;
  readonly modified: number;
}

/** A failure the HTTP layer maps to a status code. */
export class ApiError extends Error {
  readonly status: number;

  constructor(status: number, message: string) {
    super(message);
    this.status = status;
  }
}

/**
 * Parse `CELLESTRA_ROOTS` — `Name=/abs/dir;Other=/abs/dir2` — into roots. Ids are URL-safe slugs
 * of the names, so a root keeps its id across restarts as long as its name is stable.
 */
export function parseRoots(spec: string, baseDir: string): Root[] {
  const roots: Root[] = [];
  for (const part of spec.split(';')) {
    const trimmed = part.trim();
    if (!trimmed) continue;
    const eq = trimmed.indexOf('=');
    if (eq <= 0) throw new Error(`CELLESTRA_ROOTS entry '${trimmed}' must look like Name=/path`);
    const name = trimmed.slice(0, eq).trim();
    const dir = path.resolve(baseDir, trimmed.slice(eq + 1).trim());
    const id =
      name
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '-')
        .replace(/^-|-$/g, '') || `root${roots.length + 1}`;
    if (roots.some((r) => r.id === id)) throw new Error(`Duplicate root id '${id}'`);
    roots.push({ id, name, dir });
  }
  return roots;
}

/** Files the picker offers. Everything else is hidden (dirs are always listed). */
const OPENABLE = /\.(ods|csv|tsv)$/i;

/**
 * The rule `list` shows and `read`/`write` enforce: a spreadsheet file, reached without any
 * dot-named segment (dotfiles, `.git`, LibreOffice `.~lock` files, our own temp files).
 */
function checkSpreadsheetPath(rel: string | null, verb: string): void {
  const segments = checkVisiblePath(rel);
  const base = segments.at(-1) ?? '';
  if (!OPENABLE.test(base)) throw new ApiError(400, `Only .ods, .csv, and .tsv files can be ${verb}`);
}

/** Refuse a path with any dot-named segment; returns its segments. */
function checkVisiblePath(rel: string | null): string[] {
  const segments = (rel ?? '').replaceAll('\\', '/').split('/').filter(Boolean);
  if (segments.some((s) => s.startsWith('.'))) throw new ApiError(403, 'Hidden paths are not served');
  return segments;
}

export class FileApi {
  readonly roots: readonly Root[];
  /** Tail of the write queue per resolved path: check-then-write runs one at a time per file. */
  private readonly writeQueue = new Map<string, Promise<unknown>>();

  constructor(roots: readonly Root[]) {
    this.roots = roots;
  }

  root(id: string | null): Root {
    const r = this.roots.find((x) => x.id === id);
    if (!r) throw new ApiError(404, `Unknown root '${id ?? ''}'`);
    return r;
  }

  /**
   * Resolve a root-relative path inside the root, refusing anything that escapes it — `..`,
   * absolute paths, NUL bytes, and (via realpath of the parent) symlinks pointing outside.
   */
  resolve(root: Root, rel: string | null): string {
    const clean = (rel ?? '').replaceAll('\\', '/');
    if (clean.includes('\0')) throw new ApiError(400, 'Bad path');
    const full = path.resolve(root.dir, '.' + path.posix.normalize('/' + clean));
    const inside = path.relative(root.dir, full);
    if (inside.startsWith('..') || path.isAbsolute(inside)) throw new ApiError(403, 'Path escapes root');
    return full;
  }

  /** Like `resolve`, then also refuse a path whose nearest existing ancestor symlinks out of the root. */
  async resolveReal(root: Root, rel: string | null): Promise<string> {
    const full = this.resolve(root, rel);
    const rootReal = await realpath(root.dir).catch(() => root.dir);
    let probe = full;
    for (;;) {
      const real = await realpath(probe).catch(() => null);
      if (real !== null) {
        const inside = path.relative(rootReal, real);
        if (inside.startsWith('..') || path.isAbsolute(inside)) throw new ApiError(403, 'Path escapes root');
        return full;
      }
      const parent = path.dirname(probe);
      if (parent === probe) return full;
      probe = parent;
    }
  }

  async list(rootId: string | null, rel: string | null): Promise<Entry[]> {
    const root = this.root(rootId);
    checkVisiblePath(rel);
    const dir = await this.resolveReal(root, rel);
    let dirents;
    try {
      dirents = await readdir(dir, { withFileTypes: true });
    } catch (e) {
      throw fsError(e);
    }
    const entries: Entry[] = [];
    for (const d of dirents) {
      if (d.name.startsWith('.')) continue; // dotfiles, LibreOffice .~lock files, our temp files
      const isDir = d.isDirectory();
      if (!isDir && !(d.isFile() && OPENABLE.test(d.name))) continue;
      const full = path.join(dir, d.name);
      const s = await stat(full).catch(() => null);
      if (!s) continue;
      entries.push({
        name: d.name,
        path: path.relative(root.dir, full).split(path.sep).join('/'),
        kind: isDir ? 'dir' : 'file',
        size: isDir ? 0 : s.size,
        modified: s.mtimeMs,
      });
    }
    entries.sort((a, b) => (a.kind === b.kind ? a.name.localeCompare(b.name) : a.kind === 'dir' ? -1 : 1));
    return entries;
  }

  async read(
    rootId: string | null,
    rel: string | null,
  ): Promise<{ bytes: Buffer; etag: string; modified: number }> {
    const root = this.root(rootId);
    checkSpreadsheetPath(rel, 'opened');
    const full = await this.resolveReal(root, rel);
    try {
      const [bytes, s] = await Promise.all([readFile(full), stat(full)]);
      if (!s.isFile()) throw new ApiError(400, 'Not a file');
      return { bytes, etag: etagOf(bytes), modified: s.mtimeMs };
    } catch (e) {
      throw fsError(e);
    }
  }

  /**
   * Atomic save: write a temp file beside the target, fsync, rename over it. `ifMatch` is the etag
   * the client opened (conflict → 412); `'*'`-less create with `ifNoneMatch` refuses to overwrite.
   * A LibreOffice `.~lock.<name>#` file means someone has it open there (PRD §10.2) → 423.
   * Writes to one path run one at a time, so two saves from the same etag cannot both win. The
   * replacement keeps the target's mode (and owner, where permitted); a hard-linked target is
   * written in place so the link survives.
   */
  async write(
    rootId: string | null,
    rel: string | null,
    bytes: Buffer,
    opts: { ifMatch?: string | null; ifNoneMatch?: string | null; force?: boolean },
  ): Promise<{ etag: string }> {
    const root = this.root(rootId);
    checkSpreadsheetPath(rel, 'saved');
    const full = await this.resolveReal(root, rel);
    return this.serialized(await queueKey(full), () => this.writeLocked(full, bytes, opts));
  }

  /** Run `fn` after every earlier write to `key` has settled, so the If-Match check and the rename are atomic. */
  private async serialized<T>(key: string, fn: () => Promise<T>): Promise<T> {
    const prev = this.writeQueue.get(key) ?? Promise.resolve();
    const run = prev.then(fn, fn);
    const tail = run.catch(() => {});
    this.writeQueue.set(key, tail);
    try {
      return await run;
    } finally {
      if (this.writeQueue.get(key) === tail) this.writeQueue.delete(key);
    }
  }

  private async writeLocked(
    full: string,
    bytes: Buffer,
    opts: { ifMatch?: string | null; ifNoneMatch?: string | null; force?: boolean },
  ): Promise<{ etag: string }> {
    const dir = path.dirname(full);
    const base = path.basename(full);
    await mkdir(dir, { recursive: true }).catch((e: unknown) => {
      throw fsError(e);
    });

    const existing = await stat(full).catch(() => null);
    if (existing && !existing.isFile()) throw new ApiError(400, 'Not a file');
    const current = existing ? await readFile(full).catch(() => null) : null;
    if (opts.ifNoneMatch === '*' && current) throw new ApiError(412, 'A file with that name already exists');
    if (opts.ifMatch && !opts.force) {
      if (!current)
        throw new ApiError(412, 'The file was moved or deleted on the server since you opened it');
      if (opts.ifMatch !== '*' && etagOf(current) !== opts.ifMatch)
        throw new ApiError(412, 'The file changed on the server since you opened it');
    }
    if (!opts.force) {
      const lock = await stat(path.join(dir, `.~lock.${base}#`)).catch(() => null);
      if (lock) throw new ApiError(423, 'The file is open in LibreOffice');
    }

    const target = await stat(full).catch(() => null);
    try {
      if (target && target.nlink > 1) {
        // A rename would split the hard link; write through it instead (not atomic, but keeps the link).
        const handle = await open(full, fsConstants.O_WRONLY | fsConstants.O_TRUNC);
        try {
          await handle.writeFile(bytes);
          await handle.sync();
        } finally {
          await handle.close();
        }
        return { etag: etagOf(bytes) };
      }
    } catch (e) {
      throw fsError(e);
    }

    const mode = target ? target.mode & 0o7777 : 0o644;
    const tmp = path.join(dir, `.${base}.${randomBytes(6).toString('hex')}.tmp`);
    try {
      const handle = await open(tmp, fsConstants.O_CREAT | fsConstants.O_EXCL | fsConstants.O_WRONLY, mode);
      try {
        await handle.chmod(mode); // open() applies the umask; the replacement keeps the original bits
        await handle.writeFile(bytes);
        await handle.sync();
      } finally {
        await handle.close();
      }
      if (target) await chown(tmp, target.uid, target.gid).catch(() => {}); // only root may give files away
      await rename(tmp, full);
    } catch (e) {
      await unlink(tmp).catch(() => {});
      throw fsError(e);
    }
    return { etag: etagOf(bytes) };
  }
}

/**
 * The write-queue key for a path: the same file under any spelling gets the same key. Symlinks are
 * resolved (for a new file, through its nearest existing folder), and the key is case- and
 * normalization-folded, since `Book.ods` and `book.ods` are one file on a case-insensitive volume
 * (macOS, Windows, most SMB shares). On a case-sensitive volume that only serializes two distinct
 * files' saves, which costs nothing.
 */
async function queueKey(full: string): Promise<string> {
  let probe = full;
  const tail: string[] = [];
  for (;;) {
    const real = await realpath(probe).catch(() => null);
    if (real !== null)
      return path
        .join(real, ...tail)
        .normalize('NFC')
        .toLowerCase();
    const parent = path.dirname(probe);
    if (parent === probe) return full.normalize('NFC').toLowerCase();
    tail.unshift(path.basename(probe));
    probe = parent;
  }
}

export function etagOf(bytes: Uint8Array): string {
  return '"' + createHash('sha256').update(bytes).digest('base64url').slice(0, 27) + '"';
}

function fsError(e: unknown): Error {
  if (e instanceof ApiError) return e;
  const code = (e as NodeJS.ErrnoException | null)?.code;
  if (code === 'ENOENT') return new ApiError(404, 'Not found');
  if (code === 'EACCES' || code === 'EPERM') return new ApiError(403, 'Permission denied');
  if (code === 'ENOTDIR') return new ApiError(400, 'Not a folder');
  if (code === 'EISDIR') return new ApiError(400, 'Not a file');
  return e instanceof Error ? e : new Error(String(e));
}
