import { effect, untracked } from '@preact/signals';
import {
  after,
  ConflictError,
  csvLosses,
  lockWarning,
  type DocumentRef,
  type OpenedDocument,
  type WorkbookSession,
} from '@cellestra/ui';
import {
  autoSaveFailure,
  isNotFound,
  message,
  type LifecycleHost,
  type SaveOutcome,
} from './lifecycleHost.ts';
import { step, type Effect, type SaveEvent, type SaveState } from './saveChart.ts';

/** How often a file change on disk that arrived at a bad moment (a dialog, an edit) is tried again. */
const EXTERNAL_CHANGE_RETRY_MS = 1_000;

/** What the document looked like when an open began, to tell whether it changed meanwhile. */
interface OpenTicket {
  readonly seq: number;
  readonly generation: number;
  readonly edits: number;
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
  /** Where the open document is in the save chart. */
  #state: SaveState = 'clean';
  /** The document generation and edit count the chart last heard about. */
  #seenGeneration = -1;
  #seenEdits = 0;
  /** Bumped when reads are dropped: a read that started under an older number is ignored. */
  #readSeq = 0;
  /** A close waits for the write in progress, then goes ahead. */
  #closeHeld = false;

  constructor(host: LifecycleHost) {
    this.#host = host;
    // Every edit tells the chart, which restarts auto-save's countdown, so the save lands in the
    // pause after a burst of typing rather than in the middle of one.
    effect(() => {
      void host.session.version.value;
      untracked(() => this.#noticeChanges());
    });
    effect(() => {
      void host.autoSaveSeconds.value;
      untracked(() => this.#armAutoSave());
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
  #recordSaved(ref: DocumentRef, edits: number, generation: number): void {
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
    this.#fire('saved');
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
    return this.#turn(run());
  }

  /** Hold the save turn while `p` runs; a close that waited for it goes ahead afterwards. */
  async #turn(p: Promise<SaveOutcome>): Promise<SaveOutcome> {
    this.#saving = p;
    try {
      return await p;
    } finally {
      if (this.#saving === p) this.#saving = null;
      if (this.#closeHeld && !this.#saving) {
        this.#closeHeld = false;
        void this.#close();
      }
    }
  }

  /** `force` skips the changed-since-opened check; `ignoreLock` the LibreOffice lock. Each is its own answer. */
  async #save(force: boolean, ignoreLock: boolean): Promise<SaveOutcome> {
    const session = this.#session;
    const ref = session.document.peek();
    if (!ref) return this.#saveAs(session.format.peek());
    // Without a handle, a browser file (uploaded, dropped) cannot be written back: ask where.
    if (!this.#host.stores.canSaveInPlace(ref)) return this.#saveAs(session.format.peek());
    if (session.format.peek() === 'csv') {
      const keep = await this.#host.questions.csvLoss(ref.name);
      if (keep === 'cancel') return this.#gaveUp();
      if (keep === 'ods') return this.#saveAs('ods');
    }
    // Nothing unsaved: there is nothing to write.
    if (!this.#fire('save').includes('writeFile')) return this.#state === 'clean' ? 'saved' : false;
    return this.#writeInPlace(ref, { force, ignoreLock });
  }

  /**
   * Write the document to its own file (the chart said `writeFile`). A conflict asks the user
   * unless the write is auto-save's, which only stands down.
   */
  async #writeInPlace(
    ref: DocumentRef,
    options: { force: boolean; ignoreLock: boolean; background?: boolean },
  ): Promise<SaveOutcome> {
    const session = this.#session;
    const store = this.#host.stores.storeFor(ref);
    const edits = session.editCount;
    const generation = session.generation;
    const bytes = session.serialize();
    try {
      const saved = await store.save(ref, bytes, options);
      if (saved.downloaded) {
        this.#host.info(`Downloaded ${saved.name}`);
        this.#ended(generation, 'failed');
        return 'downloaded';
      }
      this.#recordSaved(saved, edits, generation);
      return 'saved';
    } catch (e) {
      if (options.background) {
        this.#autoSaveSuspendedIn = generation;
        this.#host.error(new Error(autoSaveFailure(e)));
        this.#ended(generation, 'failed');
        return false;
      }
      if (!(e instanceof ConflictError) || generation !== session.generation) {
        this.#host.error(e);
        this.#ended(generation, 'failed');
        return false;
      }
      this.#fire('versionChanged');
      const choice = await this.#host.questions.conflict(e);
      if (choice === 'overwrite') {
        this.#fire('overwrite');
        const again = e.reason === 'locked' ? { ...options, ignoreLock: true } : { ...options, force: true };
        return this.#writeInPlace(ref, again);
      }
      if (choice === 'saveAs') return this.#saveAs(session.format.peek());
      if (choice === 'reload') {
        if (this.#fire('reload').includes('readFile')) await this.#read(ref);
        return false;
      }
      this.#fire('cancel');
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
      if (keep === 'cancel') return this.#gaveUp();
      if (keep === 'ods') format = 'ods';
    }
    const name = `${base}.${format}`;
    if (isExport) return this.#export(name);
    if (!this.#fire('saveAs').includes('pickFile')) return false;
    const edits = session.editCount;
    const generation = session.generation;
    const bytes = session.serialize(format, name);
    try {
      const saved = await this.#host.stores.primary.saveAs(name, bytes);
      if (!saved) {
        this.#ended(generation, 'failed');
        return false;
      }
      if (saved.downloaded) {
        // The browser can't confirm the download was kept: keep the changes marked unsaved and
        // the recovery snapshot, so closing the tab still warns and nothing is lost.
        this.#host.info(
          `Downloaded ${saved.name}. This browser can't save in place, so it stays marked unsaved.`,
        );
        this.#ended(generation, 'failed');
        return 'downloaded';
      }
      this.#recordSaved(saved, edits, generation);
      return 'saved';
    } catch (e) {
      this.#host.error(e);
      this.#ended(generation, 'failed');
      return false;
    }
  }

  /** Export a CSV copy. It isn't the document's file, so the save chart doesn't see it. */
  async #export(name: string): Promise<SaveOutcome> {
    try {
      const saved = await this.#host.stores.primary.saveAs(name, this.#session.serialize('csv', name));
      if (!saved) return false;
      this.#host.info(`Exported ${saved.name}`);
      return 'saved';
    } catch (e) {
      this.#host.error(e);
      return false;
    }
  }

  /** A question before writing was cancelled: a conflict or close question waiting on it is answered Cancel. */
  #gaveUp(): false {
    if (this.#state === 'conflict' || this.#state === 'askingClose') this.#fire('cancel');
    return false;
  }

  /** A write or read for `generation` ended; a newer document has its own chart and ignores it. */
  #ended(generation: number, event: SaveEvent): void {
    if (generation === this.#session.generation) this.#fire(event);
  }

  // ── The save chart ────────────────────────────────────────────────────────

  /** Where the open document is in the save chart (`saveChart.ts`). */
  get saveState(): SaveState {
    return this.#state;
  }

  /**
   * Tell the chart what the session did since it last looked: another document means a new chart,
   * and a new edit count an `edit`.
   */
  #noticeChanges(): void {
    const session = this.#session;
    if (session.generation !== this.#seenGeneration) {
      this.#seenGeneration = session.generation;
      this.#seenEdits = session.editCount;
      this.#readSeq++;
      this.#state = session.modified.peek() ? 'dirty' : 'clean';
      this.#armAutoSave();
    } else if (session.editCount !== this.#seenEdits) {
      this.#seenEdits = session.editCount;
      this.#fire('edit');
    }
  }

  /**
   * Move the chart and run the effects that need nothing from the caller. Returns every effect, so
   * the caller runs the ones that do (writing, reading, asking).
   */
  #fire(event: SaveEvent): readonly Effect[] {
    const { to, effects } = step(this.#state, event);
    this.#state = to;
    for (const effect of effects) {
      switch (effect) {
        case 'startTimer':
          this.#armAutoSave();
          break;
        case 'dropRead':
          this.#readSeq++;
          break;
        case 'warnChanged':
          this.#host.info('The file changed on disk. Saving will ask before overwriting it.');
          break;
        case 'retryLater':
          this.#retryExternalChange(this.#session.generation);
          break;
        case 'closeLater':
          this.#closeHeld = true;
          break;
        case 'cancelClose':
          this.#host.native()?.cancelClose();
          break;
        // The caller runs these: they need its file, its read, or its question.
        case 'writeFile':
        case 'pickFile':
        case 'showConflict':
        case 'readFile':
        case 'takeFile':
        case 'replaceWithFile':
        case 'askClose':
        case 'closeWindow':
          break;
      }
    }
    return effects;
  }

  // ── Auto-save ─────────────────────────────────────────────────────────────

  /** (Re)start auto-save's countdown, when it is on and there is something to save. */
  #armAutoSave(): void {
    const seconds = this.#host.autoSaveSeconds.peek();
    this.#cancelAutoSave?.();
    this.#cancelAutoSave = null;
    if (seconds <= 0 || this.#autoSaveSuspended || this.#state !== 'dirty') return;
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
    if (!ref) return false;
    // A CSV that would lose formatting or other sheets asks first; one that wouldn't is fine.
    if (session.format.peek() === 'csv' && csvLosses(session.workbook.peek(), session.csvSheet()).length > 0)
      return false;
    if (!this.#host.stores.canSaveInPlace(ref)) return false;
    if (!this.#fire('timer').includes('writeFile')) return false;
    const done = await this.#turn(
      this.#writeInPlace(ref, { force: false, ignoreLock: false, background: true }),
    );
    return done === 'saved';
  }

  // ── Changes elsewhere ─────────────────────────────────────────────────────

  /**
   * Re-read the document from its store. `onlyIfClean` (a change elsewhere) is a refresh: it takes
   * the file only when nothing is unsaved and gives up when edits land while it reads. Otherwise
   * (Revert to Saved) it asks about unsaved changes unless `force`, then replaces them.
   */
  async reloadFromStore(force = false, onlyIfClean = false): Promise<void> {
    const ref = this.#session.document.peek();
    if (!ref) return;
    if (onlyIfClean) {
      if (this.#fire('refresh').includes('readFile')) await this.#read(ref);
      return;
    }
    if (!force && !(await this.confirmDiscard())) return;
    await this.#exclusive(async () => {
      const now = this.#session.document.peek();
      if (now && this.#fire('reload').includes('readFile')) await this.#read(now);
      return false;
    }, true);
  }

  /** Read the file (the chart said `readFile`) and take it if the chart still wants it. */
  async #read(ref: DocumentRef): Promise<void> {
    const session = this.#session;
    const seq = this.#readSeq;
    const generation = session.generation;
    let opened: OpenedDocument;
    try {
      opened = await this.#host.stores.storeFor(ref).reload(ref);
    } catch (e) {
      if (seq === this.#readSeq) this.#ended(generation, 'readFailed');
      this.#host.error(e);
      return;
    }
    // Dropped (an edit or Save As came first), or another document is open now.
    if (seq !== this.#readSeq || generation !== session.generation) return;
    const effects = this.#fire('readDone');
    // Same contents (a sync client rewrote it, `touch`): nothing to reload, and undo stays.
    const same = opened.ref.version !== null && opened.ref.version === ref.version;
    if (effects.includes('replaceWithFile') || (effects.includes('takeFile') && !same)) {
      this.#host.reloadInPlace(() => session.reload(opened.bytes, opened.ref));
      // What is on screen is the file now (Revert to Saved): no recovery copy to offer next time.
      void this.#host.recovery.clear();
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
    if (this.#host.dialogOpen() || this.#host.isEditing()) return this.#retryExternalChange(generation);
    return this.reloadFromStore(true, true);
  }

  #retryExternalChange(generation: number): void {
    this.#cancelExternalRetry?.();
    this.#cancelExternalRetry = after(EXTERNAL_CHANGE_RETRY_MS, () => void this.externalChange(generation));
  }

  // ── Closing ───────────────────────────────────────────────────────────────

  /** Before closing the window (Electron): save, discard, or stay. */
  async requestClose(): Promise<void> {
    if (this.#closing) return;
    this.#closing = true;
    try {
      if (!this.#host.commitOpenEdit()) this.#host.native()?.cancelClose();
      else await this.#close();
    } finally {
      this.#closing = false;
    }
  }

  /** Send `close` to the chart and carry out what it says: close, ask first, or wait for the write. */
  async #close(): Promise<void> {
    const effects = this.#fire('close');
    if (effects.includes('closeWindow')) return this.#closeWindow();
    if (!effects.includes('askClose')) return;
    const choice = await this.#host.questions.discard();
    if (choice === 'discard') {
      if (this.#fire('discard').includes('closeWindow')) await this.#closeWindow();
    } else if (choice === 'cancel') this.#fire('cancel');
    else {
      // The save holds the close and closes once it is written; a refused CSV question cancels.
      await this.save();
      if (this.#state === 'askingClose') this.#fire('cancel');
    }
  }

  async #closeWindow(): Promise<void> {
    await this.#host.recovery.clear();
    this.#host.native()?.confirmClose();
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
