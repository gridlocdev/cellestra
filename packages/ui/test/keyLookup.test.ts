import { afterEach, describe, expect, it, vi } from 'vitest';
import { CommandCatalog, type CommandSpec } from '../src/commands/command.ts';
import { KeyLookup, type KeyCapture } from '../src/commands/keyLookup.ts';
import { type KeyPress, shortcutOf } from '../src/commands/shortcut.ts';

const run = () => undefined;
const spec = (
  id: string,
  title: string,
  shortcut: string,
  extra: Partial<CommandSpec> = {},
): CommandSpec => ({
  id,
  title,
  description: '',
  group: 'File',
  shortcut,
  run,
  ...extra,
});
const catalog = () =>
  new CommandCatalog([
    spec('file.save', 'Save', 'mod+s'),
    spec('file.saveAs', 'Save As…', 'mod+shift+s'),
    spec('view.zoomIn', 'Zoom In', 'mod+=', { group: 'View', alternates: ['mod++'] }),
    spec('edit.fill', 'Fill Down', 'mod+d', { group: 'Edit' }),
  ]);

const press = (init: Partial<KeyPress> & { key: string }): KeyPress => ({
  code: '',
  shiftKey: false,
  ctrlKey: false,
  altKey: false,
  metaKey: false,
  ...init,
});

function keydown(target: EventTarget, init: Partial<KeyPress> & { key: string }): Event {
  const e = Object.assign(new Event('keydown', { cancelable: true }), press(init));
  target.dispatchEvent(e);
  return e;
}

class FakeCapture implements KeyCapture {
  on = false;
  #cb: ((key: KeyPress) => void) | null = null;
  captureKeys(on: boolean) {
    this.on = on;
  }
  onCapturedKey(cb: (key: KeyPress) => void) {
    this.#cb = cb;
    return () => {
      this.#cb = null;
    };
  }
  send(key: KeyPress) {
    this.#cb?.(key);
  }
  get subscribed() {
    return this.#cb !== null;
  }
}

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
});

describe('shortcutOf', () => {
  it('reads the chord a key press makes, and waits on a modifier alone', () => {
    expect(shortcutOf(press({ key: 's', code: 'KeyS', metaKey: true, shiftKey: true }), true)).toEqual({
      key: 's',
      mod: true,
      ctrl: false,
      alt: false,
      shift: true,
    });
    expect(shortcutOf(press({ key: 's', code: 'KeyS', ctrlKey: true }), false)?.mod).toBe(true);
    expect(shortcutOf(press({ key: 'Shift', shiftKey: true }), true)).toBeNull();
    expect(shortcutOf(press({ key: 'Meta', metaKey: true }), true)).toBeNull();
  });
});

describe('commandsForKey', () => {
  it('finds the command behind a chord, its alternates too', () => {
    const c = catalog();
    const ids = (k: KeyPress) => c.commandsForKey(k, true).map((x) => x.id);
    expect(ids(press({ key: 's', code: 'KeyS', metaKey: true }))).toEqual(['file.save']);
    expect(ids(press({ key: 'S', code: 'KeyS', metaKey: true, shiftKey: true }))).toEqual(['file.saveAs']);
    expect(ids(press({ key: '+', code: 'Equal', metaKey: true, shiftKey: true }))).toEqual(['view.zoomIn']);
    expect(ids(press({ key: 'q', code: 'KeyQ', metaKey: true }))).toEqual([]);
    expect(ids(press({ key: 's', code: 'KeyS', ctrlKey: true }))).toEqual([]);
  });
});

describe('KeyLookup', () => {
  it('looks keys up instead of running them, and Escape stops', () => {
    const win = new EventTarget();
    const lookup = new KeyLookup(catalog(), { target: win, mac: true });
    lookup.start();
    // Added after start: Node's EventTarget has no capture phase, so order stands in for it.
    const later = vi.fn();
    win.addEventListener('keydown', later);
    const e = keydown(win, { key: 'd', code: 'KeyD', metaKey: true });
    expect(e.defaultPrevented).toBe(true);
    expect(later).not.toHaveBeenCalled();
    expect(lookup.last.value?.keys).toBe('⌘D');
    expect(lookup.last.value?.commands.map((c) => c.id)).toEqual(['edit.fill']);
    expect(lookup.last.value?.gridKey).toBe(false);

    keydown(win, { key: 'ArrowDown', code: 'ArrowDown' });
    expect(lookup.last.value?.commands).toEqual([]);
    expect(lookup.last.value?.gridKey).toBe(true);

    keydown(win, { key: 'Escape', code: 'Escape' });
    expect(lookup.listening.value).toBe(false);
    keydown(win, { key: 's', code: 'KeyS', metaKey: true });
    expect(later).toHaveBeenCalledTimes(1);
    expect(lookup.last.value?.keys).toBe('↓');
  });

  it('keeps the last chord while a modifier on its own is held', () => {
    const win = new EventTarget();
    const lookup = new KeyLookup(catalog(), { target: win, mac: true });
    lookup.start();
    keydown(win, { key: 's', code: 'KeyS', metaKey: true });
    keydown(win, { key: 'Meta', metaKey: true });
    expect(lookup.last.value?.keys).toBe('⌘S');
    lookup.stop();
  });

  it('stops when the window loses focus', () => {
    const win = new EventTarget();
    const lookup = new KeyLookup(catalog(), { target: win, mac: true });
    lookup.start();
    win.dispatchEvent(new Event('blur'));
    expect(lookup.listening.value).toBe(false);
  });

  it('stops by itself after a while without keys', () => {
    vi.useFakeTimers();
    const win = new EventTarget();
    const lookup = new KeyLookup(catalog(), { target: win, idleMs: 1000, mac: true });
    lookup.start();
    vi.advanceTimersByTime(900);
    keydown(win, { key: 's', code: 'KeyS', metaKey: true });
    vi.advanceTimersByTime(900);
    expect(lookup.listening.value).toBe(true);
    vi.advanceTimersByTime(200);
    expect(lookup.listening.value).toBe(false);
  });

  it('stops listening when a lookup throws, so keys come back', () => {
    vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const c = catalog();
    vi.spyOn(c, 'commandsForKey').mockImplementation(() => {
      throw new Error('boom');
    });
    const win = new EventTarget();
    const lookup = new KeyLookup(c, { target: win, mac: true });
    lookup.start();
    keydown(win, { key: 's', code: 'KeyS', metaKey: true });
    expect(lookup.listening.value).toBe(false);
    const later = vi.fn();
    win.addEventListener('keydown', later);
    keydown(win, { key: 's', code: 'KeyS', metaKey: true });
    expect(later).toHaveBeenCalledTimes(1);
  });

  it('turns the desktop capture on and off, and looks up the keys it sends', () => {
    const capture = new FakeCapture();
    const lookup = new KeyLookup(catalog(), { target: new EventTarget(), capture, mac: true });
    lookup.start();
    expect(capture.on).toBe(true);
    capture.send(press({ key: 's', code: 'KeyS', metaKey: true }));
    expect(lookup.last.value?.commands.map((c) => c.id)).toEqual(['file.save']);
    capture.send(press({ key: 'Escape', code: 'Escape' }));
    expect(lookup.listening.value).toBe(false);
    expect(capture.on).toBe(false);
    expect(capture.subscribed).toBe(false);
  });

  it('runs every cleanup even when one fails, and stop is safe to repeat', () => {
    vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const capture = new FakeCapture();
    vi.spyOn(capture, 'captureKeys').mockImplementation((on) => {
      if (!on) throw new Error('bridge gone');
    });
    const win = new EventTarget();
    const lookup = new KeyLookup(catalog(), { target: win, capture, mac: true });
    lookup.start();
    lookup.stop();
    lookup.stop();
    expect(lookup.listening.value).toBe(false);
    expect(capture.subscribed).toBe(false);
    const later = vi.fn();
    win.addEventListener('keydown', later);
    keydown(win, { key: 'a', code: 'KeyA' });
    expect(later).toHaveBeenCalledTimes(1);
  });
});
