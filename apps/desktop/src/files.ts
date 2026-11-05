// File-system helpers for the main process. No Electron imports, so the tests can run them in Node.
import { watch, type FSWatcher } from 'node:fs';
import { chown, mkdir, open, readFile, rename, stat, unlink, writeFile } from 'node:fs/promises';
import path from 'node:path';

const SPREADSHEET = /\.(ods|csv|tsv)$/i;

export function isSpreadsheetPath(p: unknown): p is string {
  return typeof p === 'string' && p.length > 0 && !p.includes('\0') && SPREADSHEET.test(p);
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
 * Atomic save: a temp file beside the target, fsync, rename. The replacement keeps the target's
 * mode (and owner, where permitted). A hard-linked target is written in place instead, since a
 * rename would split the link.
 */
export async function writeAtomic(p: string, bytes: Uint8Array): Promise<void> {
  const target = await stat(p).catch(() => null);
  if (target && target.nlink > 1) {
    const h = await open(p, 'r+');
    try {
      await h.truncate(0);
      await h.writeFile(bytes);
      await h.sync();
    } finally {
      await h.close();
    }
    return;
  }
  const mode = target ? target.mode & 0o7777 : 0o644;
  const tmp = path.join(path.dirname(p), `.${path.basename(p)}.${process.pid}.${Date.now()}.tmp`);
  try {
    const h = await open(tmp, 'wx', mode);
    try {
      await h.chmod(mode); // open() applies the umask
      await h.writeFile(bytes);
      await h.sync();
    } finally {
      await h.close();
    }
    if (target) await chown(tmp, target.uid, target.gid).catch(() => {});
    await rename(tmp, p);
  } catch (e) {
    await unlink(tmp).catch(() => {});
    throw e;
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
  #writing: Promise<void> = Promise.resolve();

  constructor(file: string | null, limit = 200) {
    this.#file = file;
    this.#limit = limit;
  }

  async load(): Promise<void> {
    if (!this.#file) return;
    try {
      const data: unknown = JSON.parse(await readFile(this.#file, 'utf8'));
      if (!Array.isArray(data)) return;
      // Grants made while loading (argv, open-file) stay in front.
      const loaded = data.filter(isSpreadsheetPath).filter((x) => !this.#paths.includes(x));
      this.#paths = [...this.#paths, ...loaded].slice(0, this.#limit);
    } catch {
      // Missing or unreadable: start empty.
    }
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

  /** Resolves once every grant so far is on disk (for tests). */
  flushed(): Promise<void> {
    return this.#writing;
  }

  async #persist(): Promise<void> {
    if (!this.#file) return;
    try {
      await mkdir(path.dirname(this.#file), { recursive: true });
      await writeFile(this.#file, JSON.stringify(this.#paths));
    } catch {
      // Best effort: the grants still hold for this run.
    }
  }
}
