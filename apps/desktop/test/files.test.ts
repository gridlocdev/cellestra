import {
  chmod,
  link,
  lstat,
  mkdir,
  mkdtemp,
  readdir,
  readFile,
  rename,
  rm,
  stat,
  symlink,
  writeFile,
} from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import {
  changedSince,
  contentVersion,
  isAppUrl,
  PathGrants,
  pdfExportTarget,
  saveAsMismatch,
  spreadsheetArgs,
  syncDir,
  watchFile,
  withFileLock,
  writeAtomic,
} from '../src/files.ts';

let dir: string;
beforeAll(async () => {
  dir = await mkdtemp(path.join(os.tmpdir(), 'cellestra-desktop-'));
});
afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));

describe('writeAtomic', () => {
  it('keeps the file mode', async () => {
    const p = path.join(dir, 'mode.ods');
    await writeFile(p, 'a');
    await chmod(p, 0o600);
    await writeAtomic(p, new TextEncoder().encode('b'));
    expect((await stat(p)).mode & 0o777).toBe(0o600);
    expect(await readFile(p, 'utf8')).toBe('b');
  });

  it('writes through a hard link instead of splitting it', async () => {
    const p = path.join(dir, 'linked.ods');
    const other = path.join(dir, 'other-name.ods');
    await writeFile(p, 'one-longer');
    await link(p, other);
    await writeAtomic(p, new TextEncoder().encode('two'));
    expect(await readFile(other, 'utf8')).toBe('two');
    await writeAtomic(p, new TextEncoder().encode('a longer third version'));
    expect(await readFile(other, 'utf8')).toBe('a longer third version');
  });

  // root may write anything, so the permission check cannot be seen when the tests run as root.
  it.skipIf(process.getuid?.() === 0)('refuses to replace a read-only file', async () => {
    const p = path.join(dir, 'readonly.ods');
    await writeFile(p, 'keep');
    await chmod(p, 0o444);
    await expect(writeAtomic(p, new TextEncoder().encode('replaced'))).rejects.toThrow(/read-only/);
    expect(await readFile(p, 'utf8')).toBe('keep');
  });

  it('replaces the target of a symlink and keeps the link', async () => {
    const real = path.join(dir, 'real.ods');
    const alias = path.join(dir, 'alias.ods');
    await writeFile(real, 'a');
    await chmod(real, 0o600);
    await symlink(real, alias);
    await writeAtomic(alias, new TextEncoder().encode('b'));
    expect((await lstat(alias)).isSymbolicLink()).toBe(true);
    expect(await readFile(real, 'utf8')).toBe('b');
    expect((await stat(real)).mode & 0o777).toBe(0o600);
  });

  it('lets concurrent saves of one file all land, with no temp file left behind', async () => {
    const sub = path.join(dir, 'concurrent');
    await rm(sub, { recursive: true, force: true });
    await mkdir(sub);
    const p = path.join(sub, 'busy.ods');
    const bodies = Array.from({ length: 20 }, (_, i) => `v${i}`);
    await Promise.all(bodies.map((b) => writeAtomic(p, new TextEncoder().encode(b))));
    expect(bodies).toContain(await readFile(p, 'utf8'));
    expect(await readdir(sub)).toEqual(['busy.ods']);
  });

  it('runs locked work for one file one at a time, across spellings of its path', async () => {
    const order: string[] = [];
    const p = path.join(dir, 'locked.ods');
    const job = (name: string, ms: number) => async () => {
      order.push(`${name}+`);
      await wait(ms);
      order.push(`${name}-`);
    };
    await Promise.all([
      withFileLock(p, job('a', 30)),
      withFileLock(path.join(dir, '.', 'locked.ods'), job('b', 1)),
    ]);
    expect(order).toEqual(['a+', 'a-', 'b+', 'b-']);
  });

  it('fsyncs a folder, so a rename into it survives a power cut', async () => {
    await expect(syncDir(dir)).resolves.toBeUndefined();
    await expect(syncDir(path.join(dir, 'no-such-folder'))).rejects.toMatchObject({ code: 'ENOENT' });
  });

  it('creates new files', async () => {
    const p = path.join(dir, 'new.csv');
    await writeAtomic(p, new TextEncoder().encode('x'));
    expect(await readFile(p, 'utf8')).toBe('x');
  });
});

describe('watchFile', () => {
  it('keeps watching after the file is replaced by a rename', async () => {
    const p = path.join(dir, 'watched.ods');
    await writeFile(p, '0');
    let calls = 0;
    const stop = watchFile(p, () => calls++, 50);
    try {
      await wait(400); // FSEvents may still report the initial write
      for (const n of [1, 2]) {
        const before = calls;
        const tmp = path.join(dir, `.tmp-${n}`);
        await writeFile(tmp, String(n));
        await rename(tmp, p);
        await wait(400);
        expect(calls).toBe(before + 1); // debounced to one call, and still alive after the replace
      }
      const before = calls;
      await writeFile(path.join(dir, 'unrelated.ods'), 'x');
      await wait(400);
      expect(calls).toBe(before);
    } finally {
      stop();
    }
  });
});

describe('main-process guards', () => {
  it('accepts only the app origin', () => {
    expect(isAppUrl('app://cellestra/index.html')).toBe(true);
    expect(isAppUrl('app://cellestra.evil/index.html')).toBe(false);
    expect(isAppUrl('https://example.com/')).toBe(false);
    expect(isAppUrl(undefined)).toBe(false);
  });

  it('resolves argv against the directory it ran in', () => {
    expect(
      spreadsheetArgs(['/Apps/Cellestra', '--flag', 'a.ods', 'notes.txt', '/abs/b.csv'], '/home/me'),
    ).toEqual(['/home/me/a.ods', '/abs/b.csv']);
  });

  it('grants only spreadsheet paths, and remembers them across runs', async () => {
    const file = path.join(dir, 'grants.json');
    const g = new PathGrants(file);
    await g.load();
    g.grant('/x/a.ods');
    g.grant('/x/.env');
    expect(g.has('/x/a.ods')).toBe(true);
    expect(g.has('/x/b.ods')).toBe(false);
    expect(g.has('/x/.env')).toBe(false);
    expect(g.has('relative.ods')).toBe(false);
    await g.flushed();
    const again = new PathGrants(file);
    await again.load();
    expect(again.has('/x/a.ods')).toBe(true);
  });

  it('keeps saved grants when a file is granted before they load', async () => {
    const file = path.join(dir, 'grants-early.json');
    await writeFile(file, JSON.stringify(['/x/old.ods']));
    const g = new PathGrants(file);
    g.grant('/x/argv.ods'); // argv / open-file arrive before app ready
    await wait(20);
    expect(JSON.parse(await readFile(file, 'utf8'))).toEqual(['/x/old.ods']);
    await g.load();
    await g.flushed();
    expect(JSON.parse(await readFile(file, 'utf8'))).toEqual(['/x/argv.ods', '/x/old.ods']);
  });

  it('never writes the grants file from an instance that lost the single-instance lock', async () => {
    const file = path.join(dir, 'grants-second.json');
    await writeFile(file, JSON.stringify(['/x/kept.ods']));
    const g = new PathGrants(file);
    g.disablePersist();
    await g.load();
    g.grant('/x/new.ods');
    await g.flushed();
    expect(g.has('/x/new.ods')).toBe(true);
    expect(JSON.parse(await readFile(file, 'utf8'))).toEqual(['/x/kept.ods']);
  });
});

describe('saveAsMismatch', () => {
  it('accepts the extension the bytes were prepared for, in any case', () => {
    expect(saveAsMismatch('Budget.ods', '/tmp/Other.ODS')).toBeNull();
    expect(saveAsMismatch('Budget.csv', '/tmp/x.csv')).toBeNull();
  });

  it('refuses a different spreadsheet extension, so .ods bytes never land in a .csv', () => {
    expect(saveAsMismatch('Budget.ods', '/tmp/Budget.csv')).toMatch(/\.ods extension/);
    expect(saveAsMismatch('Budget.csv', '/tmp/Budget.ods')).toMatch(/\.csv extension/);
  });

  it('refuses a path that is not a spreadsheet', () => {
    expect(saveAsMismatch('Budget.ods', '/tmp/Budget.txt')).toMatch(/Save as/);
  });
});

describe('pdfExportTarget', () => {
  it('asks before overwriting when .pdf was added to the chosen name and that file exists', async () => {
    await writeFile(path.join(dir, 'report.pdf'), 'old');
    expect(await pdfExportTarget(path.join(dir, 'report'))).toEqual({
      path: path.join(dir, 'report.pdf'),
      askOverwrite: true,
    });
  });

  it('leaves the question to the save dialog when the chosen name already ends in .pdf', async () => {
    await writeFile(path.join(dir, 'kept.PDF'), 'old');
    expect(await pdfExportTarget(path.join(dir, 'kept.PDF'))).toEqual({
      path: path.join(dir, 'kept.PDF'),
      askOverwrite: false,
    });
  });

  it('does not ask for a new file', async () => {
    expect(await pdfExportTarget(path.join(dir, 'fresh'))).toEqual({
      path: path.join(dir, 'fresh.pdf'),
      askOverwrite: false,
    });
  });
});

describe('changedSince', () => {
  const bytes = (s: string) => new TextEncoder().encode(s);

  it('sees a same-size edit that kept the mtime', () => {
    const opened = contentVersion(bytes('1,2,3'));
    expect(changedSince(opened, { bytes: bytes('1,2,3'), mtimeMs: 5000 })).toBe(false);
    expect(changedSince(opened, { bytes: bytes('1,2,4'), mtimeMs: 5000 })).toBe(true);
  });

  it('ignores a touch that changed only the mtime', () => {
    const opened = contentVersion(bytes('same'));
    expect(changedSince(opened, { bytes: bytes('same'), mtimeMs: 9999 })).toBe(false);
  });

  it('still understands an mtime-only version from an older session', () => {
    expect(changedSince('5000', { bytes: bytes('x'), mtimeMs: 5000.4 })).toBe(false);
    expect(changedSince('5000', { bytes: bytes('x'), mtimeMs: 7000 })).toBe(true);
  });
});
