// File-system helpers for the main process. No Electron imports, so the tests can run them in Node.
import { constants as fsConstants, watch, type FSWatcher } from 'node:fs';
import { createHash, randomBytes } from 'node:crypto';
import {
  access,
  chown,
  mkdir,
  open,
  readFile,
  readlink,
  realpath,
  rename,
  stat,
  unlink,
} from 'node:fs/promises';
import path from 'node:path';

const SPREADSHEET = /\.(ods|csv|tsv)$/i;

export function isSpreadsheetPath(p: unknown): p is string {
  return typeof p === 'string' && p.length > 0 && !p.includes('\0') && SPREADSHEET.test(p);
}

/** The lower-case extension of a spreadsheet path ('ods', 'csv' or 'tsv'), or null. */
export function spreadsheetFormat(p: string): 'ods' | 'csv' | 'tsv' | null {
  const m = SPREADSHEET.exec(p);
  return m ? (m[1]!.toLowerCase() as 'ods' | 'csv' | 'tsv') : null;
}

/**
 * Save As gets bytes the renderer already serialized for `suggested`'s format, so a chosen path with
 * another extension would hold the wrong bytes (zip data in a .csv). Returns why to refuse, or null.
 */
export function saveAsMismatch(suggested: string, chosen: string): string | null {
  const want = spreadsheetFormat(suggested);
  const got = spreadsheetFormat(chosen);
  if (!got) return 'Save as .ods, .csv, or .tsv';
  if (want && got !== want) return `This copy was prepared as .${want}; save it with the .${want} extension.`;
  return null;
}

/**
 * Where Export PDF writes for a name chosen in the save dialog. The dialog already asked before
 * overwriting the name it returned, but when that name lacks `.pdf` we add it, and the dialog never
 * saw that file: `askOverwrite` says it exists and the user must be asked first.
 */
export async function pdfExportTarget(chosen: string): Promise<{ path: string; askOverwrite: boolean }> {
  if (chosen.toLowerCase().endsWith('.pdf')) return { path: chosen, askOverwrite: false };
  const target = `${chosen}.pdf`;
  const exists = await stat(target).then(
    () => true,
    () => false,
  );
  return { path: target, askOverwrite: exists };
}

const HASH_PREFIX = 'sha256:';
/** mtimes within this many ms count as equal (some file systems round them). */
const MTIME_SLACK_MS = 1;

/** The version of a file's contents: a hash, so an edit that keeps the size and mtime still shows. */
export function contentVersion(bytes: Uint8Array): string {
  return HASH_PREFIX + createHash('sha256').update(bytes).digest('base64url');
}

/**
 * Did the file change since it was read at `expected` (a `contentVersion`)? A touch that keeps the
 * contents is not a change. A bare number is an mtime version from an older session's recovery data,
 * and is compared as an mtime.
 */
export function changedSince(expected: string, current: { bytes: Uint8Array; mtimeMs: number }): boolean {
  if (expected.startsWith(HASH_PREFIX)) return contentVersion(current.bytes) !== expected;
  const mtime = Number(expected);
  return !Number.isFinite(mtime) || Math.abs(current.mtimeMs - mtime) > MTIME_SLACK_MS;
}

/** Only the app's own pages may use the bridge (not a navigated-away or injected frame). */
export function isAppUrl(url: string | undefined | null): boolean {
  return typeof url === 'string' && url.startsWith('app://cellestra/');
}

/** Spreadsheet paths from a command line, made absolute against the directory it ran in. */
export function spreadsheetArgs(argv: readonly string[], cwd: string): string[] {
  return argv.filter((a) => !a.startsWith('-') && SPREADSHEET.test(a)).map((a) => path.resolve(cwd, a));
}

/**
 * Atomic save: a temp file beside the target, fsync, rename. A symlink is followed, so the link
 * stays and its target is replaced. The replacement keeps the target's mode (and owner, where
 * permitted). A hard-linked target is written in place instead, since a rename would split the link.
 * A file we may not write (read-only, mode 0444) is refused, not replaced, even though the folder
 * would allow the rename.
 * Each call uses its own random temp name and only ever removes that one, so two saves of one file
 * cannot delete each other's temp; `withFileLock` orders them.
 */
export async function writeAtomic(p: string, bytes: Uint8Array): Promise<void> {
  const real = await realTarget(path.resolve(p));
  const target = await stat(real).catch(() => null);
  if (target && !(await writable(real))) throw new Error(`“${path.basename(p)}” is read-only`);
  if (target && target.nlink > 1) {
    // Not atomic: write the new bytes over the old ones first and cut to length only after, so a
    // crash part-way leaves the new bytes at the front (possibly with old ones after), never an
    // empty file.
    const h = await open(real, 'r+');
    try {
      await h.writeFile(bytes);
      await h.truncate(bytes.length);
      await h.sync();
    } finally {
      await h.close();
    }
    return;
  }
  const mode = target ? target.mode & 0o7777 : 0o644;
  const tmp = path.join(path.dirname(real), `.${path.basename(real)}.${randomBytes(8).toString('hex')}.tmp`);
  let created = false;
  try {
    const h = await open(tmp, 'wx', mode);
    created = true;
    try {
      await h.chmod(mode); // open() applies the umask
      await h.writeFile(bytes);
      await h.sync();
    } finally {
      await h.close();
    }
    if (target) await chown(tmp, target.uid, target.gid).catch(() => {});
    await rename(tmp, real);
  } catch (e) {
    if (created) await unlink(tmp).catch(() => {});
    throw e;
  }
}

async function writable(p: string): Promise<boolean> {
  try {
    await access(p, fsConstants.W_OK);
    return true;
  } catch {
    return false;
  }
}

/** The file a path names after following symlinks; a path that doesn't exist yet is its own target. */
async function realTarget(abs: string): Promise<string> {
  const real = await realpath(abs).catch(() => null);
  if (real !== null) return real;
  const link = await readlink(abs).catch(() => null); // dangling: write where it points
  return link === null ? abs : path.resolve(path.dirname(abs), link);
}

const fileLocks = new Map<string, Promise<unknown>>();

/**
 * Run `fn` after every earlier `withFileLock` call for the same file has settled, so a save's
 * check-then-write cannot interleave with another save of that file. Paths are compared resolved,
 * and case-folded where the volume usually is case-insensitive (macOS, Windows).
 */
export async function withFileLock<T>(p: string, fn: () => Promise<T>): Promise<T> {
  const abs = path.resolve(p);
  const key = process.platform === 'linux' ? abs : abs.normalize('NFC').toLowerCase();
  const prev = fileLocks.get(key) ?? Promise.resolve();
  const run = prev.then(fn, fn);
  const tail = run.catch(() => {});
  fileLocks.set(key, tail);
  try {
    return await run;
  } finally {
    if (fileLocks.get(key) === tail) fileLocks.delete(key);
  }
}

/**
 * Watch one file by watching its folder, so the watch survives the file being replaced by a
 * rename (an atomic save, ours or another app's). Events for other names are ignored; a burst
 * of events becomes one call after `delayMs`.
 */
export function watchFile(p: string, onChange: () => void, delayMs = 300): () => void {
  const dir = path.dirname(p);
  const base = path.basename(p);
  let timer: NodeJS.Timeout | null = null;
  let watcher: FSWatcher | null = null;
  try {
    watcher = watch(dir, (_event, filename) => {
      // Some platforms omit the name; then any change in the folder counts.
      if (filename !== null && filename !== undefined && filename.toString() !== base) return;
      if (timer) clearTimeout(timer);
      timer = setTimeout(() => {
        timer = null;
        onChange();
      }, delayMs);
    });
    watcher.on('error', () => {});
  } catch {
    watcher = null;
  }
  return () => {
    watcher?.close();
    watcher = null;
    if (timer) clearTimeout(timer);
    timer = null;
  };
}

/**
 * The files the renderer may read, save and watch: ones the user picked in a dialog, opened from
 * the OS (open-file, argv), or opened before (persisted, so recents and recovery still work).
 */
export class PathGrants {
  readonly #file: string | null;
  readonly #limit: number;
  #paths: string[] = [];
  /** Persisting waits for `load()`, so a grant made first (argv, open-file) can't overwrite the file. */
  #writing: Promise<void>;
  #loaded: () => void = () => {};
  #persistEnabled = true;

  constructor(file: string | null, limit = 200) {
    this.#file = file;
    this.#limit = limit;
    this.#writing = new Promise((resolve) => (this.#loaded = resolve));
  }

  async load(): Promise<void> {
    try {
      if (!this.#file) return;
      const data: unknown = JSON.parse(await readFile(this.#file, 'utf8'));
      if (!Array.isArray(data)) return;
      // Grants made while loading (argv, open-file) stay in front.
      const loaded = data.filter(isSpreadsheetPath).filter((x) => !this.#paths.includes(x));
      this.#paths = [...this.#paths, ...loaded].slice(0, this.#limit);
    } catch {
      // Missing or unreadable: start empty.
    } finally {
      this.#loaded();
    }
  }

  /**
   * Keep grants in memory only, never on disk: for an instance that lost the single-instance lock
   * and is about to quit, so it can't race the running app's writes.
   */
  disablePersist(): void {
    this.#persistEnabled = false;
  }

  has(p: unknown): p is string {
    return isSpreadsheetPath(p) && path.isAbsolute(p) && this.#paths.includes(path.resolve(p));
  }

  grant(p: string): void {
    if (!isSpreadsheetPath(p)) return;
    const abs = path.resolve(p);
    this.#paths = [abs, ...this.#paths.filter((x) => x !== abs)].slice(0, this.#limit);
    this.#writing = this.#writing.then(() => this.#persist());
  }

  /** Resolves once every grant so far is on disk (for tests). Waits for `load()` first. */
  flushed(): Promise<void> {
    return this.#writing;
  }

  async #persist(): Promise<void> {
    if (!this.#file || !this.#persistEnabled) return;
    try {
      await mkdir(path.dirname(this.#file), { recursive: true });
      await writeAtomic(this.#file, new TextEncoder().encode(JSON.stringify(this.#paths)));
    } catch {
      // Best effort: the grants still hold for this run.
    }
  }
}
