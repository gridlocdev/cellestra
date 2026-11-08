import { existsSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs';
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
import { CliError } from './errors.ts';
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
    throw new CliError(
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
  fillFromCsv(book, book.activeSheet, text, { delimiter: csv?.delimiter ?? ',', locale });
  return { workbook: book, package: null };
}

export function recalculate(workbook: Workbook): RecalculationResult {
  const engine = new RecalculationEngine(workbook);
  engine.rebuildGraph();
  return engine.recalculateAll();
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

/** Write next to the target, then rename over it: a reader never sees half a file. */
export function atomicWrite(path: string, bytes: Uint8Array): void {
  const full = resolve(path);
  const temp = resolve(dirname(full), `.${basename(full)}.${randomUUID().replaceAll('-', '')}.tmp`);
  try {
    writeFileSync(temp, bytes);
    renameSync(temp, full);
  } catch (e) {
    rmSync(temp, { force: true });
    throw e;
  }
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
