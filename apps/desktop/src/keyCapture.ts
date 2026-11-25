import type { CapturedKey } from './bridge.ts';

/** The fields of Electron's `Input` that the key lookup reads. */
export interface KeyInput {
  readonly type: string;
  readonly key: string;
  readonly code: string;
  readonly shift: boolean;
  readonly control: boolean;
  readonly alt: boolean;
  readonly meta: boolean;
}

/**
 * What the window does with a key while the Keyboard Shortcuts lookup captures keys:
 * - `pass`: the key goes on as usual (key-ups, and everything when not capturing).
 * - `release`: Escape. Capturing ends here, in the main process, and the page gets the Escape
 *   too, so the keyboard comes back even when the page has stopped answering.
 * - `forward`: held back from the menu bar and the page (so ⌘Q is looked up, not run) and sent
 *   to the page as a `CapturedKey`.
 */
export function keyCaptureAction(capturing: boolean, input: KeyInput): 'pass' | 'release' | 'forward' {
  if (!capturing || input.type === 'keyUp') return 'pass';
  if (input.key === 'Escape') return 'release';
  return 'forward';
}

export function capturedKey(input: KeyInput): CapturedKey {
  return {
    key: input.key,
    code: input.code,
    shiftKey: input.shift,
    ctrlKey: input.control,
    altKey: input.alt,
    metaKey: input.meta,
  };
}
