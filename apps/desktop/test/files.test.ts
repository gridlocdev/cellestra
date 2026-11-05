import { chmod, link, mkdtemp, readFile, rename, rm, stat, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { isAppUrl, PathGrants, spreadsheetArgs, watchFile, writeAtomic } from '../src/files.ts';

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
});
