import { beforeAll, describe, expect, it, vi } from 'vitest';
import type { NativeMenuItem } from '@cellestra/ui';

// The controller expects a browser; give it just enough of one.
beforeAll(() => {
  vi.stubGlobal('localStorage', { getItem: () => null, setItem: () => {}, removeItem: () => {} });
  vi.stubGlobal('matchMedia', () => ({ matches: false, addEventListener: () => {} }));
  vi.stubGlobal('document', { documentElement: { dataset: {} } });
});

describe('the native menu', () => {
  it('is sent only when what it shows changed', async () => {
    const { AppController } = await import('../src/app/controller.ts');
    const { buildCommands } = await import('../src/app/commands.ts');
    const { nativeMenu, menuSender } = await import('../src/app/nativeMenu.ts');
    const app = new AppController();
    const catalog = buildCommands(app, true);
    const sent: (readonly NativeMenuItem[])[] = [];
    const menu = menuSender((items) => sent.push(items));
    const send = (items: readonly NativeMenuItem[]) => menu.send(items);

    expect(send(nativeMenu(catalog))).toBe(true);
    // The cursor moves but nothing the menu shows changes: Electron isn't asked to rebuild it.
    app.selection.select(3, 2);
    expect(send(nativeMenu(catalog))).toBe(false);
    expect(sent).toHaveLength(1);

    // Undo becomes possible: that is a change, and the new menu goes out.
    app.session.execute(
      (await import('@cellestra/core')).EditCommands.setInput('Sheet1', 0, 0, '1', 'en-US'),
    );
    expect(send(nativeMenu(catalog))).toBe(true);
    expect(sent).toHaveLength(2);
    // Another window had the menu bar: the same menu goes out again.
    menu.forget();
    expect(send(nativeMenu(catalog))).toBe(true);
    const edit = sent[1]!.find((m) => m.label === 'Edit')!;
    expect(edit.submenu!.find((m) => m.command === 'edit.undo')?.enabled).toBe(true);
  });
});

describe('watchMenuInputs', () => {
  it('works the menu out again when any state a menu item shows changes', async () => {
    const { AppController } = await import('../src/app/controller.ts');
    const { watchMenuInputs } = await import('../src/app/nativeMenu.ts');
    const app = new AppController();
    const calls: boolean[] = [];
    const stop = watchMenuInputs(app, (soon) => calls.push(soon));
    calls.length = 0;
    const changes: [string, () => void][] = [
      ['layout', () => (app.layout.value = app.layout.value === 'ribbon' ? 'bar' : 'ribbon')],
      ['ribbon collapsed', () => (app.ribbonCollapsed.value = !app.ribbonCollapsed.value)],
      ['format painter', () => (app.formatPainter.value = {} as never)],
      ['find query', () => (app.findQuery.value = 'x')],
      ['NAS available', () => (app.nasAvailable.value = !app.nasAvailable.value)],
      ['show formulas', () => (app.showFormulas.value = !app.showFormulas.value)],
    ];
    for (const [what, change] of changes) {
      change();
      expect(calls, what).toEqual([true]);
      calls.length = 0;
    }
    app.selection.select(3, 2);
    expect(calls).toEqual([false]);
    stop();
  });
});

describe('menuSchedule', () => {
  it('waits for the cursor to stand still, but sends an edit at once', async () => {
    const { menuSchedule } = await import('../src/app/nativeMenu.ts');
    vi.useFakeTimers();
    try {
      let built = 0;
      const changed = menuSchedule(() => built++, 100);
      for (let i = 0; i < 10; i++) {
        changed(false); // an arrow key held down
        vi.advanceTimersByTime(30);
      }
      expect(built).toBe(0);
      vi.advanceTimersByTime(100);
      expect(built).toBe(1);
      // Enter commits an edit and moves the cursor: the edit doesn't wait for the cursor.
      changed(true);
      changed(false);
      vi.advanceTimersByTime(0);
      expect(built).toBe(2);
    } finally {
      vi.useRealTimers();
    }
  });
});
