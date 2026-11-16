import { ConflictError, type DocumentRef, type WorkbookSession } from '@cellestra/ui';
import type { ReadonlySignal } from '@preact/signals';
import type { DocumentStores } from './documentStores.ts';

/** How a save ended: written, handed to the browser as a download, or not done. */
export type SaveOutcome = 'saved' | 'downloaded' | false;

/** The Electron calls the lifecycle makes: which file to watch, and the answer to a close request. */
export interface LifecycleNative {
  watch(path: string | null): void;
  confirmClose(): void;
  cancelClose(): void;
}

/** The questions the lifecycle asks the user, through the app's own dialogs. */
export interface LifecycleQuestions {
  /** Unsaved changes are about to be replaced: save them, discard them, or stay. */
  discard(): Promise<'save' | 'discard' | 'cancel'>;
  /**
   * Saving `name` as CSV keeps one sheet's values: keep CSV, save as .ods instead, or stay.
   * Resolves 'csv' without asking when nothing would be lost.
   */
  csvLoss(name: string): Promise<'csv' | 'ods' | 'cancel'>;
  /** The file changed or is locked elsewhere. */
  conflict(e: ConflictError): Promise<'overwrite' | 'saveAs' | 'reload' | 'cancel'>;
}

/** Everything outside the lifecycle it needs: the session, the stores, the UI's answers and notices. */
export interface LifecycleHost {
  readonly session: WorkbookSession;
  readonly stores: DocumentStores;
  /** Seconds between an edit and its auto-save; 0 turns auto-save off. */
  readonly autoSaveSeconds: ReadonlySignal<number>;
  /** Electron's bridge; null in a browser. Read at each call. */
  native(): LifecycleNative | null;
  readonly questions: LifecycleQuestions;
  /** Commit the open cell edit. False when the value was refused and the editor stays open. */
  commitOpenEdit(): boolean;
  /** Whether a cell is being edited. */
  isEditing(): boolean;
  /** Whether a dialog is open. */
  dialogOpen(): boolean;
  info(text: string): void;
  /** A problem worth an error toast: an exception, or a warning's text. */
  error(e: unknown): void;
  /** Crash-recovery storage. */
  readonly recovery: {
    clear(): Promise<void>;
    keepAside(title: string, ref: DocumentRef | null, bytes: Uint8Array): Promise<void>;
  };
  /** Add a file to the recent files. */
  remember(ref: DocumentRef): Promise<void>;
  /** Another document is on screen (opened or new): reset per-document UI and focus the grid. */
  replaced(): void;
  /** Reload the same document's new bytes, keeping the cursor where it was. */
  reloadInPlace(run: () => void): void;
}

/** An error's text, for a toast. */
export function message(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}

/** The file is gone (not just unreadable right now): browser NotFoundError, NAS 404, native ENOENT. */
export function isNotFound(e: unknown): boolean {
  if (!(e instanceof Error) && !(e instanceof DOMException)) return false;
  return e.name === 'NotFoundError' || /\bENOENT\b/.test(e.message);
}

/** What the one message for a failed auto-save says: why it stopped and what restarts it. */
export function autoSaveFailure(e: unknown): string {
  const resume = 'Auto-save is paused until you save with ⌘S.';
  if (e instanceof ConflictError) {
    const why =
      e.reason === 'locked' ? 'the file is locked by another app' : 'the file was changed by another app';
    return `Auto-save stopped: ${why}. ${resume}`;
  }
  if (e instanceof DOMException && e.name === 'NotAllowedError')
    return `Auto-save needs your permission to write the file. ${resume}`;
  return `Auto-save couldn't write the file: ${message(e)}. ${resume}`;
}
