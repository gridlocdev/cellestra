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

  it('matches literal ctrl shortcuts to Ctrl on every platform', () => {
    const insertFunction = parseShortcut('ctrl+alt+f');
    const ctrlAltF = key({ key: 'f', code: 'KeyF', ctrlKey: true, altKey: true });
    // Windows and Linux: Ctrl is both the mod key and the literal ctrl key.
    expect(matches(insertFunction, ctrlAltF, false)).toBe(true);
    expect(matches(insertFunction, key({ key: 'f', code: 'KeyF', altKey: true }), false)).toBe(false);
    expect(matches(parseShortcut('ctrl+pagedown'), key({ key: 'PageDown', ctrlKey: true }), false)).toBe(
      true,
    );
    expect(matches(parseShortcut('ctrl+pageup'), key({ key: 'PageUp', ctrlKey: true }), false)).toBe(true);
    expect(matches(parseShortcut('ctrl+pageup'), key({ key: 'PageUp' }), false)).toBe(false);
    // macOS: ctrl is ⌃, separate from ⌘.
    expect(matches(insertFunction, ctrlAltF, true)).toBe(true);
    expect(matches(insertFunction, key({ key: 'f', code: 'KeyF', metaKey: true, altKey: true }), true)).toBe(
      false,
    );
    expect(matches(parseShortcut('mod+k'), key({ key: 'k', code: 'KeyK', ctrlKey: true }), true)).toBe(false);
  });

  it('labels literal ctrl shortcuts as Ctrl outside macOS', () => {
    expect(formatShortcut(parseShortcut('ctrl+alt+f'), false)).toBe('Ctrl+Alt+F');
    expect(formatShortcut(parseShortcut('ctrl+pagedown'), false)).toBe('Ctrl+Pagedown');
    expect(formatShortcut(parseShortcut('ctrl+alt+f'), true)).toBe('⌃⌥F');
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
    expect(catalog.search('sa').map((c) => c.id)).toEqual(['file.save', 'data.sort']);
  });

  it('ignores spaces and matches letters in order', () => {
    expect(catalog.search('sortasc').map((c) => c.id)).toEqual(['data.sort']);
    expect(catalog.search('s ort  asc').map((c) => c.id)).toEqual(['data.sort']);
    expect(catalog.search('srtasc').map((c) => c.id)).toEqual(['data.sort']);
    expect(catalog.search('bd').map((c) => c.id)).toEqual(['format.bold']);
    expect(catalog.search('xyz')).toEqual([]);
  });

  it('only global commands fire while typing', () => {
    const e = key({ key: 'b', code: 'KeyB', ...mod });
    expect(catalog.forEvent(e, true)).toBeUndefined();
    expect(catalog.forEvent(key({ key: 's', code: 'KeyS', ...mod }), true)?.id).toBe('file.save');
  });
});

describe('CommandCatalog.forEvent', () => {
  it('lets an exact shifted shortcut beat a punctuation alternate listed earlier', () => {
    const run = vi.fn();
    const catalog = new CommandCatalog([
      {
        id: 'view.zoomIn',
        title: 'Zoom In',
        description: '',
        group: 'View',
        shortcut: 'mod+=',
        alternates: ['mod++'],
        run,
      },
      {
        id: 'insert.rows',
        title: 'Insert Rows',
        description: '',
        group: 'Insert',
        shortcut: 'mod+shift+=',
        run,
      },
    ]);
    // ⇧⌘= on a US layout: the key is "=", the character "+".
    const shiftEquals = key({ key: '+', code: 'Equal', shiftKey: true, ...mod });
    expect(catalog.forEvent(shiftEquals, false)?.id).toBe('insert.rows');
    // A layout with its own "+" key, unshifted, still zooms in.
    const plus = key({ key: '+', code: 'NumpadAdd', ...mod });
    expect(catalog.forEvent(plus, false)?.id).toBe('view.zoomIn');
  });
});
