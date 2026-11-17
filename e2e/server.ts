// Where the scenarios and the driver find the app: the NAS server serving the built client over fixture data.
import { execFileSync } from 'node:child_process';
import { cpSync, mkdirSync, readdirSync, rmSync, utimesSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

export const ROOT = fileURLToPath(new URL('..', import.meta.url));
export const DATA = `${ROOT}.e2e-data`;
// Its own port, apart from the app's default (8787) and from other apps' checks, so they can run at the same time.
export const PORT = 8791;
export const BASE_URL = `http://127.0.0.1:${PORT}`;
/** The page's clock in every scenario and in the driver, so dates and "Saved" times read the same on every run. */
export const NOW = new Date('2027-01-15T09:00:00Z');
/** Every fixture file's modified time, so the Open dialog's dates look the same in every run and screenshot. */
const FIXTURE_TIME = new Date(NOW.getTime() - 2 * 24 * 60 * 60 * 1000);

export const SERVER_COMMAND = ['apps/server/src/main.ts'];
export const SERVER_ENV: Record<string, string> = {
  CELLESTRA_PORT: String(PORT),
  CELLESTRA_HOST: '127.0.0.1',
  CELLESTRA_ROOTS: `E2E=${DATA}`,
  CELLESTRA_STATIC: `${ROOT}apps/web/dist`,
};

/** Replace .e2e-data with a fresh copy of e2e/fixtures/<fixture>. Never point this at real data. */
export function seedData(fixture: string): void {
  rmSync(DATA, { recursive: true, force: true });
  mkdirSync(DATA, { recursive: true });
  cpSync(fileURLToPath(new URL(`./fixtures/${fixture}`, import.meta.url)), DATA, { recursive: true });
  for (const file of readdirSync(DATA, { recursive: true, encoding: 'utf8' })) {
    utimesSync(`${DATA}/${file}`, FIXTURE_TIME, FIXTURE_TIME);
  }
}

const CLI = `${ROOT}apps/cli/src/main.ts`;

/** A cell of a workbook under .e2e-data as the disk has it now (read with the CLI), to check the file and not just the screen. */
export function readCell(file: string, cell: string): string {
  return execFileSync(process.execPath, [CLI, 'get', `${DATA}/${file}`, cell], {
    encoding: 'utf8',
  }).replace(/\n$/, '');
}

/** Change a workbook under .e2e-data the way another program would, e.g. `writeCells('budget.ods', 'A1=Hello')`. */
export function writeCells(file: string, ...edits: string[]): void {
  execFileSync(process.execPath, [CLI, 'set', `${DATA}/${file}`, '--in-place', ...edits]);
}
