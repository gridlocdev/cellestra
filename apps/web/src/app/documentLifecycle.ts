import { effect, type ReadonlySignal } from '@preact/signals';
import {
  after,
  ConflictError,
  csvLosses,
  lockWarning,
  type DocumentRef,
  type OpenedDocument,
  type WorkbookSession,
} from '@cellestra/ui';
import type { DocumentStores } from './documentStores.ts';

/** How a save ended: written, handed to the browser as a download, or not done. */
export type SaveOutcome = 'saved' | 'downloaded' | false;

/** How often a file change on disk that arrived at a bad moment (a dialog, an edit) is tried again. */
const EXTERNAL_CHANGE_RETRY_MS = 1_000;

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

/** What the document looked like when an open began, to tell whether it changed meanwhile. */
interface OpenTicket {
  readonly seq: number;
  readonly generation: number;
  readonly edits: number;
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

/**
 * The open document's life: opening (a newer open wins, edits made meanwhile are asked about),
 * saving one at a time, auto-save, reloading after a change on disk, and closing. It knows nothing
 * about how dialogs are drawn; the host answers its questions.
 */
export class DocumentLifecycle {
  readonly #host: LifecycleHost;
  /** The save (or Save As) in flight; a second one waits for it instead of racing it. */
  #saving: Promise<SaveOutcome> | null = null;
  /** A close request is showing its prompt; repeated requests don't stack dialogs. */
  #closing = false;
  #cancelAutoSave: (() => void) | null = null;
  /**
   * The document generation in which an auto-save write failed: auto-save stays down for that
   * document until a save of any kind succeeds. Opening another document starts afresh.
   */
  #autoSaveSuspendedIn: number | null = null;
  /** Bumped by every open and New: an open whose number is no longer current lost the race. */
  #openSeq = 0;
  /** A file change on disk waiting for a better moment to be looked at. */
  #cancelExternalRetry: (() => void) | null = null;

  constructor(host: LifecycleHost) {
    this.#host = host;
    // Every edit (and every change of interval) restarts auto-save's countdown, so the save lands
    // in the pause after a burst of typing rather than in the middle of one.
    effect(() => {
      void host.session.version.value;
      void host.autoSaveSeconds.value;
      this.#armAutoSave();
    });
  }

  get #session(): WorkbookSession {
    return this.#host.session;
  }

  /** Whether a save is running. */
  get saving(): boolean {
    return this.#saving !== null;
  }

  /**
   * A bad moment to replace the document unasked (a file changed elsewhere): a save is running, a
   * dialog is open, or a cell is being edited.
   */
  get busy(): boolean {
    return this.#saving !== null || this.#host.dialogOpen() || this.#host.isEditing();
  }

  /** Changes that closing the tab would lose: unsaved edits, or a cell edit not committed yet. */
  get hasUnsavedWork(): boolean {
    return this.#session.modified.peek() || this.#host.isEditing();
  }

  // ── Opening ───────────────────────────────────────────────────────────────

  /** Start a blank document, after asking about unsaved changes. */
  async newDocument(): Promise<void> {
    if (!(await this.confirmDiscard())) return;
    this.#openSeq++; // an open still on its way must not replace the new document
    this.#session.newDocument();
    this.#host.replaced();
    void this.#host.recovery.clear();
  }

  /**
   * Open what `fetch` brings (null when the user cancelled), after asking about unsaved changes.
   * A failure is shown; when it says the file is gone, `gone` runs first (a recent file leaves the
   * list).
   */
  async open(fetch: () => Promise<OpenedDocument | null>, gone?: () => Promise<void>): Promise<void> {
    const ticket = await this.#startOpen();
    if (!ticket) return;
    try {
      const opened = await fetch();
      if (opened) await this.#finishOpen(opened, ticket);
    } catch (e) {
      // Only a file that is gone leaves the list; a denied permission or an offline NAS doesn't.
      if (gone && isNotFound(e)) await gone();
      this.#host.error(e);
    }
  }

  /** Ask about unsaved changes, then note what the document looked like when the open began. */
  async #startOpen(): Promise<OpenTicket | null> {
    if (!(await this.confirmDiscard())) return null;
    return { seq: ++this.#openSeq, generation: this.#session.generation, edits: this.#session.editCount };
  }

  /**
   * Adopt a document that took a while to arrive. A newer open or New wins outright; edits made
   * while it downloaded are asked about again instead of being replaced unasked.
   */
  async #finishOpen(opened: OpenedDocument, ticket: OpenTicket): Promise<void> {
    if (ticket.seq !== this.#openSeq) return;
    const changed =
      this.#session.generation !== ticket.generation || this.#session.editCount !== ticket.edits;
    if (changed && (!(await this.confirmDiscard()) || ticket.seq !== this.#openSeq)) return;
    this.adopt(opened);
  }

  /** Make `opened` the document on screen. */
  adopt(opened: OpenedDocument): void {
    try {
      this.#session.load(opened.bytes, opened.ref);
    } catch (e) {
      return this.#host.error(new Error(`${opened.ref.name} could not be opened: ${message(e)}`));
    }
    void this.#host.recovery.clear();
    void this.#host.remember(opened.ref);
    this.#host.native()?.watch(opened.ref.store === 'native' ? opened.ref.id : null);
    // Only a warning: the lock is LibreOffice's, and saving still asks before overwriting.
    if (opened.lock) this.#host.error(lockWarning(opened.lock));
    this.#host.replaced();
  }

  /** True when it is safe to replace the document (saved, discarded, or clean). */
  async confirmDiscard(): Promise<boolean> {
    if (!this.#host.commitOpenEdit()) return false;
    if (!this.#session.modified.peek()) return true;
    const choice = await this.#host.questions.discard();
    if (choice === 'cancel') return false;
    if (choice === 'discard') return true;
    const done = await this.#exclusive(() => this.#save(false, false));
    // Only downloaded: the document is about to be replaced and its snapshot cleared, but the
    // browser can't say the download was kept, so a copy waits for the next start.
    if (done === 'downloaded')
      await this.#host.recovery.keepAside(
        this.#session.title.peek(),
        this.#session.document.peek(),
        this.#session.serialize('ods'),
      );
    return done !== false;
  }

  // ── Saving ────────────────────────────────────────────────────────────────

  /** Save in place when the store can; otherwise (new document, browser without a handle) Save As. */
  save(force = false): Promise<boolean> {
    return this.#exclusive(() => this.#save(force, false)).then((done) => done !== false);
  }

  saveAs(format: 'ods' | 'csv' = this.#session.format.peek()): Promise<boolean> {
    // Save As and exports always run: a background save that left the document clean doesn't
    // answer the user's request for a new file.
    return this.#exclusive(() => this.#saveAs(format), true).then((done) => done !== false);
  }

  /**
   * Record a save. `edits` is the session's edit count when the bytes were serialized, and
   * `generation` the session's document generation then: when another document has been opened
   * since, the file is only remembered, not made the open document's.
   */
  saved(ref: DocumentRef, edits = this.#session.editCount, generation = this.#session.generation): void {
    if (generation !== this.#session.generation) {
      void this.#host.remember(ref);
      this.#host.info(`Saved ${ref.name}`);
      return;
    }
    this.#session.markSaved(ref, edits);
    this.#autoSaveSuspendedIn = null;
    if (!this.#session.modified.peek()) void this.#host.recovery.clear();
    void this.#host.remember(ref);
    this.#host.native()?.watch(ref.store === 'native' ? ref.id : null);
    this.#host.info(`Saved ${ref.name}`);
  }

  /**
   * One save at a time. A save asked for while another runs waits for it, then runs only if there
   * is still something to save (so a second ⌘S never races the first into a false conflict), or
   * when `always` (Save As, exports) asks for it regardless.
   */
  async #exclusive(run: () => Promise<SaveOutcome>, always = false): Promise<SaveOutcome> {
    while (this.#saving) {
      const result = await this.#saving.catch((): SaveOutcome => false);
      if (!always && !this.#session.modified.peek()) return result;
    }
    const p = run();
    this.#saving = p;
    try {
      return await p;
    } finally {
      if (this.#saving === p) this.#saving = null;
    }
  }

  /** `force` skips the changed-since-opened check; `ignoreLock` the LibreOffice lock. Each is its own answer. */
  async #save(force: boolean, ignoreLock: boolean): Promise<SaveOutcome> {
    const session = this.#session;
    const ref = session.document.peek();
    if (!ref) return this.#saveAs(session.format.peek());
    // Without a handle, a browser file (uploaded, dropped) cannot be written back: ask where.
    if (!this.#host.stores.canSaveInPlace(ref)) return this.#saveAs(session.format.peek());
    const store = this.#host.stores.storeFor(ref);
    if (session.format.peek() === 'csv') {
      const keep = await this.#host.questions.csvLoss(ref.name);
      if (keep === 'cancel') return false;
      if (keep === 'ods') return this.#saveAs('ods');
    }
    const edits = session.editCount;
    const generation = session.generation;
    const bytes = session.serialize();
    try {
      const saved = await store.save(ref, bytes, { force, ignoreLock });
      if (saved.downloaded) {
        this.#host.info(`Downloaded ${saved.name}`);
        return 'downloaded';
      }
      this.saved(saved, edits, generation);
      return 'saved';
    } catch (e) {
      if (!(e instanceof ConflictError)) {
        this.#host.error(e);
        return false;
      }
      const choice = await this.#host.questions.conflict(e);
      if (choice === 'overwrite')
        return e.reason === 'locked' ? this.#save(force, true) : this.#save(true, ignoreLock);
      if (choice === 'saveAs') return this.#saveAs(session.format.peek());
      if (choice === 'reload') await this.reloadFromStore(true);
      return false;
    }
  }

  async #saveAs(format: 'ods' | 'csv'): Promise<SaveOutcome> {
    const session = this.#session;
    const base = session.title.peek().replace(/\.[^.]+$/, '');
    // A CSV export of an .ods document is a copy: the document stays the .ods.
    const isExport = format === 'csv' && session.format.peek() === 'ods';
    if (format === 'csv' && !isExport) {
      const keep = await this.#host.questions.csvLoss(session.title.peek());
      if (keep === 'cancel') return false;
      if (keep === 'ods') format = 'ods';
    }
    const name = `${base}.${format}`;
    const edits = session.editCount;
    const generation = session.generation;
    const bytes = session.serialize(format, name);
    try {
      const saved = await this.#host.stores.primary.saveAs(name, bytes);
      if (!saved) return false;
      if (isExport) {
        this.#host.info(`Exported ${saved.name}`);
        return 'saved';
      }
      if (saved.downloaded) {
        // The browser can't confirm the download was kept: keep the changes marked unsaved and
        // the recovery snapshot, so closing the tab still warns and nothing is lost.
        this.#host.info(
          `Downloaded ${saved.name}. This browser can't save in place, so it stays marked unsaved.`,
        );
        return 'downloaded';
      }
      this.saved(saved, edits, generation);
      return 'saved';
    } catch (e) {
      this.#host.error(e);
      return false;
    }
  }

  // ── Auto-save ─────────────────────────────────────────────────────────────

  /** (Re)start auto-save's countdown, when it is on and there is something to save. */
  #armAutoSave(): void {
    const seconds = this.#host.autoSaveSeconds.peek();
    this.#cancelAutoSave?.();
    this.#cancelAutoSave = null;
    if (seconds <= 0 || this.#autoSaveSuspended || !this.#session.modified.peek()) return;
    this.#cancelAutoSave = after(seconds * 1000, () => void this.autoSave());
  }

  get #autoSaveSuspended(): boolean {
    return this.#autoSaveSuspendedIn === this.#session.generation;
  }

  /**
   * The auto-save tick: write the document in place, if it can be written without a question.
   * Never opens a dialog — a new document, a browser file without a handle, a CSV that would lose
   * formatting, and a file changed on disk all wait for a ⌘S instead. A failed write (a refused
   * permission and a conflict included) says so once and stands auto-save down until a save
   * succeeds. Returns whether it wrote.
   */
  async autoSave(): Promise<boolean> {
    const session = this.#session;
    this.#cancelAutoSave?.();
    this.#cancelAutoSave = null;
    if (this.#host.autoSaveSeconds.peek() <= 0 || this.#autoSaveSuspended) return false;
    if (this.#saving) {
      // A save is running: try again once it is done, in case edits landed meanwhile.
      void this.#saving.catch(() => false).finally(() => this.#armAutoSave());
      return false;
    }
    const ref = session.document.peek();
    if (!ref || !session.modified.peek()) return false;
    // A CSV that would lose formatting or other sheets asks first; one that wouldn't is fine.
    if (session.format.peek() === 'csv' && csvLosses(session.workbook.peek(), session.csvSheet()).length > 0)
      return false;
    if (!this.#host.stores.canSaveInPlace(ref)) return false;
    const store = this.#host.stores.storeFor(ref);
    const run = async (): Promise<SaveOutcome> => {
      const edits = session.editCount;
      const generation = session.generation;
      const bytes = session.serialize();
      try {
        this.saved(await store.save(ref, bytes, { force: false, background: true }), edits, generation);
        return 'saved';
      } catch (e) {
        this.#autoSaveSuspendedIn = generation;
        this.#host.error(new Error(autoSaveFailure(e)));
        return false;
      }
    };
    const p = run();
    this.#saving = p;
    try {
      return (await p) !== false;
    } finally {
      if (this.#saving === p) this.#saving = null;
    }
  }

  // ── Changes elsewhere ─────────────────────────────────────────────────────

  /**
   * Re-read the document from its store. `onlyIfClean` (external changes) gives up when edits
   * landed while the file was being read, instead of throwing them away.
   */
  async reloadFromStore(force = false, onlyIfClean = false): Promise<void> {
    const session = this.#session;
    const ref = session.document.peek();
    if (!ref) return;
    if (!force && !(await this.confirmDiscard())) return;
    const edits = session.editCount;
    const generation = session.generation;
    try {
      const opened = await this.#host.stores.storeFor(ref).reload(ref);
      // Another document was opened meanwhile: this reload is for a file no longer on screen.
      if (session.generation !== generation) return;
      // Same contents (a sync client rewrote it, `touch`): nothing to reload, and undo stays.
      if (onlyIfClean && opened.ref.version !== null && opened.ref.version === ref.version) return;
      if (onlyIfClean && (session.editCount !== edits || session.document.peek() !== ref)) {
        this.#host.info('The file changed on disk. Saving will ask before overwriting it.');
        return;
      }
      this.#host.reloadInPlace(() => session.reload(opened.bytes, opened.ref));
    } catch (e) {
      this.#host.error(e);
    }
  }

  /**
   * Electron noticed the file change on disk: reload silently when clean (unless the contents are
   * the same), warn when dirty. While a save, a dialog or a cell edit is open it waits and tries
   * again; repeat events meanwhile fold into one.
   */
  async externalChange(generation = this.#session.generation): Promise<void> {
    this.#cancelExternalRetry?.();
    this.#cancelExternalRetry = null;
    if (generation !== this.#session.generation) return; // another document is open now
    if (this.busy) {
      this.#cancelExternalRetry = after(EXTERNAL_CHANGE_RETRY_MS, () => void this.externalChange(generation));
      return;
    }
    if (!this.#session.modified.peek()) return this.reloadFromStore(true, true);
    this.#host.info('The file changed on disk. Saving will ask before overwriting it.');
  }

  // ── Closing ───────────────────────────────────────────────────────────────

  /** Before closing the window (Electron): save, discard, or stay. */
  async requestClose(): Promise<void> {
    if (this.#closing) return;
    this.#closing = true;
    try {
      if (await this.confirmDiscard()) {
        await this.#host.recovery.clear();
        this.#host.native()?.confirmClose();
      } else this.#host.native()?.cancelClose();
    } finally {
      this.#closing = false;
    }
  }

  /**
   * Close the document (⌘W): ask about unsaved changes, then close the window (desktop) or go
   * back to a blank, untitled workbook (browser), which is where the app starts.
   */
  async closeDocument(): Promise<void> {
    if (this.#host.native()) return this.requestClose();
    await this.newDocument();
  }
}
