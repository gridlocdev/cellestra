import { signal } from '@preact/signals';
import { after } from '../schedule.ts';
import type { Command, CommandCatalog } from './command.ts';
import { formatShortcut, isMac, type KeyPress, shortcutOf } from './shortcut.ts';

/** What the last key pressed during a lookup does. */
export interface KeyLookupResult {
  /** The keys as the app writes them (⇧⌘L, Ctrl+Shift+L). */
  readonly keys: string;
  readonly commands: readonly Command[];
  /** No ⌘/Ctrl, and a key the grid itself answers: it types, moves the cursor or edits. */
  readonly gridKey: boolean;
}

const GRID_KEYS = new Set([
  'up',
  'down',
  'left',
  'right',
  'enter',
  'tab',
  'home',
  'end',
  'pageup',
  'pagedown',
  'backspace',
  'delete',
  'space',
  'f2',
]);

/**
 * The desktop app's side of a lookup. While it captures, it holds keys back from the menu bar
 * (so ⌘Q or ⌘W is looked up, not run) and passes them on. It lets Escape through and stops by
 * itself on Escape, so a page that stopped answering can't keep the keyboard.
 */
export interface KeyCapture {
  captureKeys(on: boolean): void;
  onCapturedKey(cb: (key: KeyPress) => void): () => void;
}

export interface KeyLookupOptions {
  /** Where key and focus events arrive: the window. */
  readonly target: EventTarget;
  /** The desktop app's capture, or null in a browser. */
  readonly capture?: KeyCapture | null;
  /** With no key pressed for this long, the lookup stops. */
  readonly idleMs?: number;
  readonly mac?: boolean;
}

/** Stop after this long without a key, in case the way out was missed. */
const IDLE_MS = 30_000;

/**
 * "Press a key to see what it does." While listening, every key press is looked up in the
 * catalog instead of running anything. The ways out, so the keyboard is never stuck:
 * Escape (the page and the desktop app each stop on it), the Stop button, the window losing
 * focus, closing the dialog, a while with no keys, and any error while looking a key up.
 */
export class KeyLookup {
  readonly listening = signal(false);
  readonly last = signal<KeyLookupResult | null>(null);
  readonly #catalog: CommandCatalog;
  readonly #options: KeyLookupOptions;
  #cleanups: (() => void)[] = [];
  #cancelIdle: (() => void) | null = null;

  constructor(catalog: CommandCatalog, options: KeyLookupOptions) {
    this.#catalog = catalog;
    this.#options = options;
  }

  start(): void {
    if (this.listening.peek()) return;
    this.last.value = null;
    this.listening.value = true;
    const { target, capture } = this.#options;
    // Capture phase on the window: ahead of the dialog's Escape and the app's own shortcuts.
    const onKey = (e: Event) => {
      e.preventDefault();
      e.stopImmediatePropagation();
      this.#safely(() => this.press(e as KeyboardEvent));
    };
    const onBlur = () => this.stop();
    // An options object, not `true`: some EventTargets (Node's) only match capture removal that way.
    const capturePhase = { capture: true };
    target.addEventListener('keydown', onKey, capturePhase);
    target.addEventListener('blur', onBlur);
    this.#cleanups.push(() => target.removeEventListener('keydown', onKey, capturePhase));
    this.#cleanups.push(() => target.removeEventListener('blur', onBlur));
    if (capture) {
      this.#cleanups.push(capture.onCapturedKey((key) => this.#safely(() => this.press(key))));
      capture.captureKeys(true);
      this.#cleanups.push(() => capture.captureKeys(false));
    }
    this.#restartIdle();
  }

  /** Ends the lookup. Safe to call any number of times; every cleanup runs even if one fails. */
  stop(): void {
    this.#cancelIdle?.();
    this.#cancelIdle = null;
    const cleanups = this.#cleanups;
    this.#cleanups = [];
    this.listening.value = false;
    const failures: unknown[] = [];
    for (const cleanup of cleanups.reverse()) {
      try {
        cleanup();
      } catch (e) {
        failures.push(e);
      }
    }
    if (failures.length > 0) console.error('key lookup cleanup failed', ...failures);
  }

  /** One key press while listening: Escape stops, a modifier on its own waits, anything else is looked up. */
  press(e: KeyPress): void {
    if (!this.listening.peek()) return;
    if (e.key === 'Escape') {
      this.stop();
      return;
    }
    this.#restartIdle();
    const mac = this.#options.mac ?? isMac;
    const shortcut = shortcutOf(e, mac);
    if (!shortcut) return;
    this.last.value = {
      keys: formatShortcut(shortcut, mac),
      commands: this.#catalog.commandsForKey(e, mac),
      gridKey: !shortcut.mod && !shortcut.ctrl && (shortcut.key.length === 1 || GRID_KEYS.has(shortcut.key)),
    };
  }

  #restartIdle(): void {
    this.#cancelIdle?.();
    this.#cancelIdle = after(this.#options.idleMs ?? IDLE_MS, () => this.stop());
  }

  /** A lookup that throws stops listening, so a bug can't leave every key swallowed. */
  #safely(run: () => void): void {
    try {
      run();
    } catch (e) {
      this.stop();
      console.error('key lookup failed', e);
    }
  }
}
