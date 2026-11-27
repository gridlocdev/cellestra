import { existsSync, readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import type { CommandCatalog, Shortcut } from '@cellestra/ui';

// The controller expects a browser; give it just enough of one.
beforeAll(() => {
  vi.stubGlobal('localStorage', { getItem: () => null, setItem: () => {}, removeItem: () => {} });
  vi.stubGlobal('matchMedia', () => ({ matches: false, addEventListener: () => {} }));
  vi.stubGlobal('document', { documentElement: { dataset: {} } });
});

async function catalogFor(mac: boolean): Promise<CommandCatalog> {
  const { AppController } = await import('../src/app/controller.ts');
  const { buildCommands } = await import('../src/app/commands.ts');
  return buildCommands(new AppController(), mac);
}

/** The keys a chord presses: outside macOS, `mod` and `ctrl` are the same Ctrl key. */
function chord(s: Shortcut, mac: boolean): string {
  const ctrl = mac ? `${s.mod ? 'cmd+' : ''}${s.ctrl ? 'ctrl+' : ''}` : s.mod || s.ctrl ? 'ctrl+' : '';
  return `${ctrl}${s.alt ? 'alt+' : ''}${s.shift ? 'shift+' : ''}${s.key}`;
}

describe('command catalog shortcuts', () => {
  for (const mac of [true, false]) {
    it(`gives every chord one command (${mac ? 'macOS' : 'Windows and Linux'})`, async () => {
      const catalog = await catalogFor(mac);
      const owners = new Map<string, string>();
      for (const c of catalog.all)
        for (const k of [c.shortcutKey, ...c.alternateKeys]) {
          if (!k) continue;
          const key = chord(k, mac);
          expect(owners.get(key) ?? c.id, `${key} is bound twice`).toBe(c.id);
          owners.set(key, c.id);
        }
    });
  }

  it('puts Go To on Ctrl+G and Find Next on ⌘G (macOS) or F3', async () => {
    const mac = await catalogFor(true);
    const other = await catalogFor(false);
    expect(chord(mac.get('edit.goTo')!.shortcutKey!, true)).toBe('ctrl+g');
    expect(chord(mac.get('edit.findNext')!.shortcutKey!, true)).toBe('cmd+g');
    expect(chord(other.get('edit.goTo')!.shortcutKey!, false)).toBe('ctrl+g');
    expect(chord(other.get('edit.findNext')!.shortcutKey!, false)).toBe('f3');
  });

  it('selects whole rows with Shift+Space and whole columns with Ctrl+Space', async () => {
    const catalog = await catalogFor(false);
    expect(chord(catalog.get('edit.selectRow')!.shortcutKey!, false)).toBe('shift+space');
    expect(chord(catalog.get('edit.selectColumn')!.shortcutKey!, false)).toBe('ctrl+space');
  });
});

describe('command icons', () => {
  it('gives every command on the ribbon, toolbar and right-click menu an icon that exists', async () => {
    // Without one, the button falls back to a plain square, which reads as a checkbox.
    const catalog = await catalogFor(true);
    const app = new URL('../src/app/', import.meta.url);
    const icons = new URL('../../../packages/ui/src/icons/', import.meta.url);
    const shown = new Set(
      ['Ribbon.tsx', 'Toolbar.tsx', 'ContextMenu.tsx'].flatMap((f) =>
        [...readFileSync(new URL(f, app), 'utf8').matchAll(/['"]([a-z]+\.[a-zA-Z]+)['"]/g)].map((m) => m[1]!),
      ),
    );
    const problems = catalog.all
      .filter((c) => shown.has(c.id))
      .flatMap((c) =>
        !c.icon
          ? [`${c.id}: no icon`]
          : existsSync(new URL(`${c.icon}.svg`, icons))
            ? []
            : [`${c.id}: ${c.icon}.svg missing`],
      );
    expect(problems).toEqual([]);
  });
});
