import { describe, expect, it, vi } from 'vitest';
import { CommandCatalog } from '../src/commands/command.ts';
import { formatShortcut, isMac, matches, parseShortcut, toAccelerator } from '../src/commands/shortcut.ts';

const mod = isMac ? { metaKey: true } : { ctrlKey: true };

const key = (init: Partial<KeyboardEvent> & { key: string; code?: string }) =>
  ({ metaKey: false, ctrlKey: false, shiftKey: false, altKey: false, code: '', ...init }) as KeyboardEvent;

describe('shortcuts', () => {
  it('formats per platform and for Electron', () => {
    const s = parseShortcut('mod+shift+k');
    expect(formatShortcut(s, true)).toBe('⇧⌘K');
    expect(formatShortcut(s, false)).toBe('Ctrl+Shift+K');
    expect(toAccelerator(s)).toBe('CmdOrCtrl+Shift+K');
  });

  it('matches mod to ⌘ on mac and Ctrl elsewhere, by physical key', () => {
    const s = parseShortcut('mod+k');
    expect(matches(s, key({ key: 'k', code: 'KeyK', metaKey: true }), true)).toBe(true);
    // AZERTY: the key labelled K sits elsewhere; its own letter decides.
    expect(matches(parseShortcut('mod+z'), key({ key: 'z', code: 'KeyW', metaKey: true }), true)).toBe(true);
    expect(matches(parseShortcut('mod+w'), key({ key: 'z', code: 'KeyW', metaKey: true }), true)).toBe(false);
    expect(matches(s, key({ key: 'k', code: 'KeyK', ctrlKey: true }), true)).toBe(false);
    expect(matches(s, key({ key: 'k', code: 'KeyK', ctrlKey: true }), false)).toBe(true);
    expect(
      matches(parseShortcut('mod+alt+k'), key({ key: '˚', code: 'KeyK', metaKey: true, altKey: true }), true),
    ).toBe(true);
  });

  it('lets shifted punctuation match', () => {
    expect(
      matches(parseShortcut('mod++'), key({ key: '+', code: 'Equal', metaKey: true, shiftKey: true }), true),
    ).toBe(true);
  });
});

describe('CommandCatalog', () => {
  const run = vi.fn();
  const catalog = new CommandCatalog([
    {
      id: 'file.save',
      title: 'Save',
      description: 'Save the workbook',
      group: 'File',
      shortcut: 'mod+s',
      run,
      global: true,
    },
    { id: 'format.bold', title: 'Bold', description: 'Toggle bold', group: 'Format', shortcut: 'mod+b', run },
    {
      id: 'data.sort',
      title: 'Sort Ascending',
      description: 'Sort A→Z',
      group: 'Data',
      keywords: 'order',
      run,
      enabled: () => false,
    },
  ]);

  it('runs, remembers recents, and refuses disabled commands', () => {
    expect(catalog.run('format.bold')).toBe(true);
    expect(catalog.run('data.sort')).toBe(false);
    expect(catalog.recentIds).toEqual(['format.bold']);
    expect(catalog.search('')[0]!.id).toBe('format.bold');
  });

  it('searches titles and keywords', () => {
    expect(catalog.search('ord').map((c) => c.id)).toEqual(['data.sort']);
    expect(catalog.search('sa').map((c) => c.id)).toEqual(['file.save']);
  });

  it('only global commands fire while typing', () => {
    const e = key({ key: 'b', code: 'KeyB', ...mod });
    expect(catalog.forEvent(e, true)).toBeUndefined();
    expect(catalog.forEvent(key({ key: 's', code: 'KeyS', ...mod }), true)?.id).toBe('file.save');
  });
});
