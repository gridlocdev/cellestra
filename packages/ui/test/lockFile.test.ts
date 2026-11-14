import { describe, expect, it } from 'vitest';
import { lockWarning, parseLockFile } from '../src/platform/lockFile.ts';
import { NasStore } from '../src/platform/nasStore.ts';
import { NativeStore } from '../src/platform/nativeStore.ts';
import type { CellestraNative, NativeFile } from '../src/platform/native.ts';

describe('LibreOffice lock files', () => {
  it('reads the user and the time from a lock LibreOffice wrote', () => {
    const text = 'Ada Lovelace,ada,analytical,14.11.2026 10:22,file:///home/ada/.config/libreoffice/4;';
    expect(parseLockFile(text)).toEqual({ name: 'Ada Lovelace', date: '14.11.2026 10:22' });
  });

  it('unescapes commas, semicolons and backslashes inside a field', () => {
    const owner = parseLockFile('Lovelace\\, Ada\\; Countess\\\\,ada,host,01.02.2026 08:00,;');
    expect(owner.name).toBe('Lovelace, Ada; Countess\\');
    expect(owner.date).toBe('01.02.2026 08:00');
  });

  it('falls back to the system user, then the host, when no user name was set', () => {
    expect(parseLockFile(',ada,analytical,14.11.2026 10:22,;').name).toBe('ada');
    expect(parseLockFile(',,analytical,,;')).toEqual({ name: 'analytical', date: null });
  });

  it('copes with a byte-order mark, a missing terminator and an empty file', () => {
    expect(parseLockFile('﻿Ada,ada,host,14.11.2026 10:22').name).toBe('Ada');
    expect(parseLockFile('')).toEqual({ name: null, date: null });
  });

  it('says who has the file open, and still warns when the lock names no one', () => {
    expect(lockWarning({ name: 'Ada', date: '14.11.2026 10:22' })).toBe(
      'Opened by Ada on 14.11.2026 10:22 in LibreOffice – your changes may conflict.',
    );
    expect(lockWarning({ name: 'Ada', date: null })).toBe(
      'Opened by Ada in LibreOffice – your changes may conflict.',
    );
    expect(lockWarning({ name: null, date: null })).toBe('Open in LibreOffice – your changes may conflict.');
  });
});

describe('stores report a lock on open', () => {
  const bytes = new Uint8Array([1, 2, 3]);

  it('the NAS store reads the lock the server sends with the file', async () => {
    const lock = 'Ada,ada,host,14.11.2026 10:22,;';
    const responses = [
      new Response(bytes, { headers: { ETag: '"e1"', 'X-Cellestra-Lock': encodeURIComponent(lock) } }),
      new Response(bytes, { headers: { ETag: '"e1"' } }),
    ];
    const nas = new NasStore('/api', () => Promise.resolve(responses.shift()!));
    const locked = await nas.open('docs', 'a.ods');
    expect(locked.lock).toEqual({ name: 'Ada', date: '14.11.2026 10:22' });
    expect(locked.ref.version).toBe('"e1"');
    expect((await nas.open('docs', 'a.ods')).lock).toBeUndefined();
  });

  it('the desktop store reads the lock the main process found', async () => {
    const file = (lock: string | null): NativeFile => ({
      path: '/d/a.ods',
      name: 'a.ods',
      bytes,
      modified: 0,
      version: 'h',
      lock,
    });
    let next = file('Ada,ada,host,14.11.2026 10:22,;');
    const bridge = { readFile: () => Promise.resolve(next) } as unknown as CellestraNative;
    const store = new NativeStore(bridge);
    expect((await store.openPath('/d/a.ods')).lock?.name).toBe('Ada');
    next = file(null);
    expect((await store.openPath('/d/a.ods')).lock).toBeUndefined();
  });
});
