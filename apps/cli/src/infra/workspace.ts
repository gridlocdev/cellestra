import {
  accessSync,
  chownSync,
  constants as fsConstants,
  closeSync,
  existsSync,
  fchmodSync,
  fsyncSync,
  ftruncateSync,
  openSync,
  readFileSync,
  readlinkSync,
  realpathSync,
  renameSync,
  rmSync,
  statSync,
  writeSync,
  type Stats,
} from 'node:fs';
import { randomUUID } from 'node:crypto';
import { basename, dirname, extname, resolve } from 'node:path';
import {
  fillFromCsv,
  readOdsWithPackage,
  RecalculationEngine,
  Workbook,
  writeOds,
  type CsvOptions,
  type OdsPackage,
  type RecalculationResult,
} from '@cellestra/core';
import { CliError, UsageError } from './errors.ts';
import type { Console } from './io.ts';

export type WorkbookFormat = 'ods' | 'csv';

/** A workbook plus the package it came from (null for CSV), so a save keeps unmodeled content. */
export interface LoadedWorkbook {
  readonly workbook: Workbook;
  readonly package: OdsPackage | null;
}

export const SPREADSHEET_EXTENSION = '.ods';

export function isWorkbook(path: string): boolean {
  return extname(path).toLowerCase() === SPREADSHEET_EXTENSION;
}

/** `cellestra <file>`: open an .ods in the app; anything else is probably a mistyped verb. */
export function launch(con: Console, path: string): number {
  if (!isWorkbook(path)) {
    throw new UsageError(
      `'${path}' isn't a workbook Cellestra opens (${SPREADSHEET_EXTENSION}). ` +
        'If you meant a command, `cellestra --help` lists them.',
    );
  }
  if (!existsSync(path)) throw new CliError(`'${path}' does not exist.`);
  return con.io.launch(resolve(path));
}

export function detectFormat(path: string, explicit: WorkbookFormat | null): WorkbookFormat {
  if (explicit) return explicit;
  switch (extname(path).toLowerCase()) {
    case '.ods':
      return 'ods';
    case '.csv':
    case '.tsv':
      return 'csv';
    default:
      throw new CliError(
        `Can't tell the format of '${path}' from its extension — pass --from ods|csv (or --to).`,
      );
  }
}

/** The delimiter for a CSV path: an explicit --delimiter wins, else tab for .tsv, else comma. */
export function csvDelimiter(path: string, csv: CsvOptions | null): string {
  if (csv) return csv.delimiter;
  return extname(path).toLowerCase() === '.tsv' ? '\t' : ',';
}

export function load(
  con: Console,
  path: string,
  from: WorkbookFormat | null,
  locale: string | null,
  csv: CsvOptions | null = null,
): LoadedWorkbook {
  const format = detectFormat(path, from);
  if (format === 'ods') {
    const bytes = readInputBytes(con, path);
    let loaded: { workbook: Workbook; package: OdsPackage };
    try {
      loaded = readOdsWithPackage(bytes);
    } catch (e) {
      if (e instanceof CliError) throw e;
      throw new CliError(`Could not read '${path}' as an .ods file: ${(e as Error).message}`);
    }
    if (locale !== null) loaded.workbook.settings.locale = locale;
    return { workbook: loaded.workbook, package: loaded.package };
  }
  const text = readInputText(con, path);
  const book = Workbook.createNew();
  if (locale !== null) book.settings.locale = locale;
  fillFromCsv(book, book.activeSheet, text, { delimiter: csvDelimiter(path, csv), locale });
  return { workbook: book, package: null };
}

/** Recompute every formula, whatever the calculation mode: what the app's Recalculate (F9) does. */
export function recalculate(workbook: Workbook): RecalculationResult {
  const engine = new RecalculationEngine(workbook);
  engine.rebuildGraph();
  return engine.recalculateAll();
}

/**
 * Bring results up to date after a command's edits, as the app does after an edit: everything in
 * automatic calculation, nothing in manual (results wait for `recalc`). The command's edits count
 * as one structural change, so the engine works out the whole graph again.
 */
export function recalculateAfterEdits(workbook: Workbook): RecalculationResult {
  return new RecalculationEngine(workbook).apply({ changedCells: [], changedRows: [], structural: true });
}

export function saveOds(con: Console, workbook: Workbook, pkg: OdsPackage | null, path: string): void {
  const bytes = writeOds(workbook, pkg);
  if (path === '-') {
    con.io.stdoutBytes(bytes);
    return;
  }
  atomicWrite(path, bytes);
}

export function saveText(con: Console, content: string, path: string): void {
  if (path === '-') {
    con.out(content);
    return;
  }
  atomicWrite(path, new TextEncoder().encode(content));
}

/**
 * Write next to the target, fsync, rename over it, then fsync the folder: a reader never sees half a
 * file, and the rename is on disk too. A symlink is followed, so the link stays and its target is
 * replaced. The replacement keeps the target's mode (and owner, where permitted). A read-only target
 * is refused rather than replaced. A hard-linked target is written in place instead, since a rename
 * would split the link.
 */
export function atomicWrite(path: string, bytes: Uint8Array): void {
  const full = realTarget(resolve(path));
  const target = statOrNull(full);
  if (target && !writable(full)) throw new CliError(`'${path}' is read-only; nothing was written.`);
  if (target && target.nlink > 1) {
    writeInPlace(full, bytes);
    return;
  }
  const mode = target ? target.mode & 0o7777 : 0o644;
  const temp = resolve(dirname(full), `.${basename(full)}.${randomUUID().replaceAll('-', '')}.tmp`);
  let created = false;
  try {
    const fd = openSync(temp, 'wx', mode);
    created = true;
    try {
      fchmodSync(fd, mode); // open() applies the umask; the replacement keeps the original bits
      writeAll(fd, bytes);
      fsyncSync(fd);
    } finally {
      closeSync(fd);
    }
    if (target) {
      try {
        chownSync(temp, target.uid, target.gid);
      } catch {
        // Only root may give files away.
      }
    }
    renameSync(temp, full);
  } catch (e) {
    if (created) rmSync(temp, { force: true });
    throw e;
  }
  syncDir(dirname(full));
}

/**
 * Not atomic, as the server and desktop app do it: write the new bytes over the old ones first and
 * cut to length only after, so a crash part-way leaves the new bytes at the front (possibly with old
 * ones after), never an empty file.
 */
function writeInPlace(full: string, bytes: Uint8Array): void {
  const fd = openSync(full, 'r+');
  try {
    writeAll(fd, bytes);
    ftruncateSync(fd, bytes.length);
    fsyncSync(fd);
  } finally {
    closeSync(fd);
  }
}

/** Errors from fsyncing a folder on platforms or file systems that don't allow it (Windows, some SMB). */
const DIR_SYNC_UNSUPPORTED = new Set(['EISDIR', 'EPERM', 'EINVAL', 'ENOTSUP', 'EBADF']);

/** fsync a folder, so a rename into it is on disk; where the platform can't, this does nothing. */
function syncDir(dir: string): void {
  let fd: number | null = null;
  try {
    fd = openSync(dir, 'r');
    fsyncSync(fd);
  } catch (e) {
    const code = (e as NodeJS.ErrnoException | null)?.code;
    if (code === undefined || !DIR_SYNC_UNSUPPORTED.has(code)) throw e;
  } finally {
    if (fd !== null) closeSync(fd);
  }
}

function writable(full: string): boolean {
  try {
    accessSync(full, fsConstants.W_OK);
    return true;
  } catch {
    return false;
  }
}

/** The file a path names after following symlinks; a path that doesn't exist yet is its own target. */
function realTarget(full: string): string {
  try {
    return realpathSync(full);
  } catch {
    // A dangling symlink: write where it points, so the link stays.
    try {
      const link = readlinkSync(full);
      return resolve(dirname(full), link);
    } catch {
      return full;
    }
  }
}

function statOrNull(full: string): Stats | null {
  try {
    return statSync(full);
  } catch {
    return null;
  }
}

function writeAll(fd: number, bytes: Uint8Array): void {
  let offset = 0;
  while (offset < bytes.length) offset += writeSync(fd, bytes, offset, bytes.length - offset);
}

function readInputBytes(con: Console, path: string): Uint8Array {
  if (path === '-') return con.io.readStdin();
  if (!existsSync(path)) throw new CliError(`File not found: ${path}`);
  return readFileSync(path);
}

export function readInputText(con: Console, path: string): string {
  if (path === '-') return con.stdinText();
  if (!existsSync(path)) throw new CliError(`File not found: ${path}`);
  return readFileSync(path, 'utf8');
}
