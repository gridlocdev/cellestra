import { batch, computed, effect, signal } from '@preact/signals';
import {
  AutoSumPlanner,
  clampToUsed,
  CellRange,
  CellAddress,
  CleanUp,
  ClipboardBlock,
  CompositeEditCommand,
  DataEditCommands,
  ActionCommand,
  DataRegion,
  DataScope,
  EdgeNavigator,
  EditResult,
  FilterEngine,
  InputParser,
  ValidationPreflight,
  Validator,
  CellValue,
  EditCommands,
  FillCommand,
  FormatBlock,
  HiddenLines,
  PasteBlockCommand,
  SheetDimensions,
  StructuralEditCommand,
  parseCsv,
  type CellStyle,
  type DataValidation,
  type EditCommand,
  type FilterSpec,
  type Hyperlink,
  type NamedRange,
  type SortSpec,
  type Table,
  type Worksheet,
} from '@cellestra/core';
import {
  BrowserStore,
  ConflictError,
  NasStore,
  NativeStore,
  Selection,
  WorkbookSession,
  clearRecovery,
  csvLosses,
  darkPalette,
  forgetRecent,
  lightPalette,
  nativeBridge,
  orphanedRecoveries,
  recentFiles,
  rememberRecent,
  saveRecovery,
  type CellestraNative,
  type DocumentRef,
  type FileStore,
  type GridController,
  type GridRange,
  type OpenedDocument,
  type RecentEntry,
} from '@cellestra/ui';

import type { ChartDefinition } from '@cellestra/charts';

export type Theme = 'system' | 'light' | 'dark';

/** The one modal at a time. Promise-backed dialogs resolve through `resolve`. */
export type DialogState =
  | { readonly kind: 'confirm-discard'; readonly resolve: (choice: 'save' | 'discard' | 'cancel') => void }
  | {
      readonly kind: 'conflict';
      readonly reason: ConflictError['reason'];
      readonly message: string;
      readonly resolve: (choice: 'overwrite' | 'saveAs' | 'reload' | 'cancel') => void;
    }
  | {
      readonly kind: 'nas';
      readonly mode: 'open' | 'save';
      readonly suggestedName?: string;
      readonly resolve: (r: { root: string; path: string } | null) => void;
    }
  | {
      readonly kind: 'prompt';
      readonly title: string;
      readonly label: string;
      readonly value: string;
      readonly resolve: (v: string | null) => void;
    }
  | { readonly kind: 'find' }
  | { readonly kind: 'recovery'; readonly name: string; readonly resolve: (restore: boolean) => void }
  | {
      readonly kind: 'csv-loss';
      readonly name: string;
      readonly losses: readonly string[];
      readonly resolve: (choice: 'csv' | 'ods' | 'cancel') => void;
    }
  | { readonly kind: 'shortcuts' }
  | { readonly kind: 'about' }
  | { readonly kind: 'settings' }
  | { readonly kind: 'formatCells'; readonly styles: ReturnType<AppController['selectionStyles']> }
  | {
      readonly kind: 'filter';
      readonly target: DataTarget;
      readonly column: number;
      readonly anchor: { x: number; y: number; w: number; h: number };
    }
  | {
      readonly kind: 'validationList';
      readonly rule: DataValidation;
      readonly row: number;
      readonly column: number;
      readonly anchor: { x: number; y: number; w: number; h: number };
    }
  | { readonly kind: 'sortFilter'; readonly target: DataTarget; readonly sortOnly: boolean }
  | { readonly kind: 'validation'; readonly range: CellRange; readonly existing: DataValidation | null }
  | { readonly kind: 'hyperlink'; readonly row: number; readonly column: number }
  | { readonly kind: 'namedRanges' }
  | { readonly kind: 'functions' };

/** What the data commands act on: a table, else the current region (or selection). */
export interface DataTarget {
  readonly range: CellRange;
  readonly hasHeader: boolean;
  readonly table: Table | null;
}

const RECOVERY_INTERVAL_MS = 30_000;
const THEME_KEY = 'cellestra.theme';

/**
 * Everything the app does, independent of how it is drawn: documents and stores, edits on the
 * selection, clipboard, dialogs. Commands (`commands.ts`) and components call into this.
 */
export class AppController {
  readonly session = new WorkbookSession();
  readonly selection = new Selection();
  readonly grid: { current: GridController | null } = { current: null };

  readonly native: CellestraNative | null = nativeBridge();
  readonly browser = new BrowserStore();
  readonly nativeStore: NativeStore | null = this.native ? new NativeStore(this.native) : null;
  readonly nas = new NasStore();
  /** Whether the NAS server answered `/api/info` (the web client is being served by it). */
  readonly nasAvailable = signal(false);

  readonly dialog = signal<DialogState | null>(null);
  readonly toast = signal<{ text: string; tone: 'info' | 'error' } | null>(null);
  readonly paletteOpen = signal(false);
  readonly sheetPickerOpen = signal(false);
  readonly recentOpen = signal(false);
  readonly recents = signal<RecentEntry[]>([]);
  readonly zoom = signal(1);
  readonly showFormulas = signal(false);
  readonly theme = signal<Theme>((localStorage.getItem(THEME_KEY) as Theme | null) ?? 'system');
  readonly prefersDark = signal(matchMedia('(prefers-color-scheme: dark)').matches);
  readonly isDark = computed(() =>
    this.theme.value === 'system' ? this.prefersDark.value : this.theme.value === 'dark',
  );
  readonly palette = computed(() => (this.isDark.value ? darkPalette : lightPalette));
  readonly formatPainter = signal<FormatBlock | null>(null);
  readonly findQuery = signal('');

  #clipboard: { block: ClipboardBlock; text: string } | null = null;
  /** Set while reloading the same document, so the cursor stays put. */
  #keepSelection = false;
  #toastTimer: ReturnType<typeof setTimeout> | null = null;
  /** The session version the last recovery snapshot was taken at. */
  #snapshotVersion = -1;
  /** The save (or Save As) in flight; a second one waits for it instead of racing it. */
  #saving: Promise<boolean> | null = null;
  /** A close request is showing its prompt; repeated requests don't stack dialogs. */
  #closing = false;
  /** Settles when `start()` has finished (recovery offered): OS file opens wait for it. */
  readonly #ready: Promise<void>;
  #markReady: () => void = () => {};

  constructor() {
    this.#ready = new Promise<void>((resolve) => (this.#markReady = resolve));
    this.session.captureSelection = () => this.selection.state.peek();
    this.session.restoreSelection = (s) => (this.selection.state.value = s);
    matchMedia('(prefers-color-scheme: dark)').addEventListener(
      'change',
      (e) => (this.prefersDark.value = e.matches),
    );
    effect(() => {
      localStorage.setItem(THEME_KEY, this.theme.value);
      document.documentElement.dataset.theme = this.isDark.value ? 'dark' : 'light';
    });
    // Switching sheets starts at A1 (Calc keeps a per-sheet cursor; good enough for now).
    let lastSheet = this.session.activeSheet.peek();
    effect(() => {
      const sheet = this.session.activeSheet.value;
      if (sheet !== lastSheet) {
        lastSheet = sheet;
        if (this.#keepSelection) return;
        this.selection.select(0, 0);
      }
    });
  }

  /** Detect the NAS server, wire Electron, offer crash recovery, open a file passed by the OS. */
  async start(): Promise<void> {
    try {
      if (!this.native) {
        this.nasAvailable.value = (await this.nas.info()) !== null;
        addEventListener('beforeunload', (e) => {
          if (this.session.modified.peek()) e.preventDefault();
        });
      }
      void this.refreshRecents();
      setInterval(() => void this.snapshot(), RECOVERY_INTERVAL_MS);
      addEventListener('focus', () => this.session.recalculateVolatile());
      await this.offerRecovery();
    } finally {
      this.#markReady();
    }
  }

  /**
   * Offer the snapshots no open tab owns (each tab keeps its own), newest first. Restoring one
   * stops the round; the rest wait for the next start. Discarded ones are deleted.
   */
  async offerRecovery(): Promise<void> {
    const orphans = await orphanedRecoveries().catch(() => []);
    for (const snapshot of orphans) {
      const restore = await new Promise<boolean>(
        (resolve) => (this.dialog.value = { kind: 'recovery', name: snapshot.name, resolve }),
      );
      this.dialog.value = null;
      if (!restore) {
        await clearRecovery(snapshot.key);
        continue;
      }
      try {
        // Snapshots are always .ods bytes. A CSV document keeps saving as CSV, so its ref stays valid.
        const csv = /\.(csv|tsv)$/i.test(snapshot.ref?.name ?? snapshot.name);
        this.session.load(
          snapshot.bytes,
          snapshot.ref,
          snapshot.name.replace(/\.(csv|tsv)$/i, '.ods'),
          csv ? 'csv' : 'ods',
        );
        this.session.touch();
        await this.snapshot(); // now held under this tab's key
        await clearRecovery(snapshot.key);
      } catch (e) {
        this.error(e);
      }
      return;
    }
  }

  // ── Stores & documents ───────────────────────────────────────────────────

  storeFor(ref: DocumentRef): FileStore {
    if (ref.store === 'native' && this.nativeStore) return this.nativeStore;
    if (ref.store === 'nas') return this.nas;
    return this.browser;
  }

  /** Where Open/Save As go by default: Electron's file system, else the NAS, else this computer. */
  get primaryStore(): 'native' | 'nas' | 'browser' {
    return this.nativeStore ? 'native' : this.nasAvailable.peek() ? 'nas' : 'browser';
  }

  async newDocument(): Promise<void> {
    if (!(await this.confirmDiscard())) return;
    this.session.newDocument();
    void clearRecovery();
    this.grid.current?.focus();
  }

  async open(): Promise<void> {
    if (!(await this.confirmDiscard())) return;
    const kind = this.primaryStore;
    let opened: OpenedDocument | null;
    try {
      if (kind === 'native') opened = await this.nativeStore!.pick();
      else if (kind === 'nas') {
        const where = await this.askNas('open');
        opened = where ? await this.nas.open(where.root, where.path) : null;
      } else opened = await this.browser.pick();
    } catch (e) {
      return this.error(e);
    }
    if (opened) this.adopt(opened);
  }

  /** Upload from this computer even when the NAS is the primary store. */
  async openLocal(): Promise<void> {
    if (!(await this.confirmDiscard())) return;
    try {
      const opened = await this.browser.pick();
      if (opened) this.adopt(opened);
    } catch (e) {
      this.error(e);
    }
  }

  /** Open a dropped file; with a File System Access handle it can be saved back in place. */
  async openFile(file: File, handle?: Promise<FileSystemHandle | null>): Promise<void> {
    if (!(await this.confirmDiscard())) return;
    try {
      const h = await handle?.catch(() => null);
      this.adopt(
        await this.browser.openFile(file, h?.kind === 'file' ? (h as FileSystemFileHandle) : undefined),
      );
    } catch (e) {
      this.error(e);
    }
  }

  async openNativePath(path: string): Promise<void> {
    // A file passed at launch waits for the recovery prompt, so neither replaces the other unasked.
    await this.#ready;
    if (!this.nativeStore || !(await this.confirmDiscard())) return;
    try {
      this.adopt(await this.nativeStore.openPath(path));
    } catch (e) {
      this.error(e);
    }
  }

  async openRecent(entry: RecentEntry): Promise<void> {
    if (!(await this.confirmDiscard())) return;
    try {
      const ref = entry.ref;
      const opened =
        ref.store === 'nas'
          ? await this.nas.open(...splitNasId(ref.id))
          : ref.store === 'native' && this.nativeStore
            ? await this.nativeStore.openPath(ref.id)
            : await this.browser.reload(ref);
      this.adopt(opened);
    } catch (e) {
      // Only a file that is gone leaves the list; a denied permission or an offline NAS doesn't.
      if (isNotFound(e)) {
        await forgetRecent(entry);
        void this.refreshRecents();
      }
      this.error(e);
    }
  }

  adopt(opened: OpenedDocument): void {
    try {
      this.session.load(opened.bytes, opened.ref);
    } catch (e) {
      return this.error(new Error(`${opened.ref.name} could not be opened: ${message(e)}`));
    }
    void clearRecovery();
    void this.remember(opened.ref);
    this.native?.watch(opened.ref.store === 'native' ? opened.ref.id : null);
    this.grid.current?.focus();
  }

  /** Save in place when the store can; otherwise (new document, browser without a handle) Save As. */
  save(force = false): Promise<boolean> {
    return this.#exclusive(() => this.#save(force));
  }

  saveAs(format: 'ods' | 'csv' = this.session.format.peek()): Promise<boolean> {
    return this.#exclusive(() => this.#saveAs(format));
  }

  /**
   * One save at a time. A save asked for while another runs waits for it, then runs only if there
   * is still something to save (so a second ⌘S never races the first into a false conflict).
   */
  async #exclusive(run: () => Promise<boolean>): Promise<boolean> {
    while (this.#saving) {
      const result = await this.#saving.catch(() => false);
      if (!this.session.modified.peek()) return result;
    }
    const p = run();
    this.#saving = p;
    try {
      return await p;
    } finally {
      if (this.#saving === p) this.#saving = null;
    }
  }

  async #save(force: boolean): Promise<boolean> {
    const ref = this.session.document.peek();
    if (!ref) return this.#saveAs(this.session.format.peek());
    const store = this.storeFor(ref);
    // Without a handle, a browser file (uploaded, dropped) cannot be written back: ask where.
    if (!store.capabilities.saveInPlace || (ref.store === 'browser' && !ref.handle))
      return this.#saveAs(this.session.format.peek());
    if (this.session.format.peek() === 'csv') {
      const keep = await this.askCsvLoss(ref.name);
      if (keep === 'cancel') return false;
      if (keep === 'ods') return this.#saveAs('ods');
    }
    const edits = this.session.editCount;
    const bytes = this.session.serialize();
    try {
      const saved = await store.save(ref, bytes, { force });
      this.saved(saved, edits);
      return true;
    } catch (e) {
      if (!(e instanceof ConflictError)) {
        this.error(e);
        return false;
      }
      const choice = await this.askConflict(e);
      if (choice === 'overwrite') return this.#save(true);
      if (choice === 'saveAs') return this.#saveAs(this.session.format.peek());
      if (choice === 'reload') await this.reloadFromStore(true);
      return false;
    }
  }

  async #saveAs(format: 'ods' | 'csv'): Promise<boolean> {
    const base = this.session.title.peek().replace(/\.[^.]+$/, '');
    // A CSV export of an .ods document is a copy: the document stays the .ods.
    const isExport = format === 'csv' && this.session.format.peek() === 'ods';
    if (format === 'csv' && !isExport) {
      const keep = await this.askCsvLoss(this.session.title.peek());
      if (keep === 'cancel') return false;
      if (keep === 'ods') format = 'ods';
    }
    const name = `${base}.${format}`;
    const edits = this.session.editCount;
    const bytes = this.session.serialize(format, name);
    try {
      let saved: DocumentRef | null;
      const kind = this.primaryStore;
      if (kind === 'native') saved = await this.nativeStore!.saveAs(name, bytes);
      else if (kind === 'nas') saved = await this.saveToNas(name, bytes);
      else saved = await this.browser.saveAs(name, bytes);
      if (!saved) return false;
      if (isExport) {
        this.info(`Exported ${saved.name}`);
        return true;
      }
      this.saved(saved, edits);
      return true;
    } catch (e) {
      this.error(e);
      return false;
    }
  }

  /**
   * Saving a CSV document keeps one sheet's values. When that loses something (other sheets,
   * formulas, formatting), ask first. Resolves 'csv' straight away when nothing would be lost.
   */
  async askCsvLoss(name: string): Promise<'csv' | 'ods' | 'cancel'> {
    const losses = csvLosses(this.session.workbook.peek(), this.session.activeSheet.peek());
    if (losses.length === 0) return 'csv';
    const choice = await new Promise<'csv' | 'ods' | 'cancel'>(
      (resolve) => (this.dialog.value = { kind: 'csv-loss', name, losses, resolve }),
    );
    this.dialog.value = null;
    return choice;
  }

  /** Save to this computer (download / picker) even when the NAS is primary. */
  async downloadCopy(): Promise<void> {
    const name = this.session.title.peek().replace(/\.[^.]+$/, '') + '.ods';
    try {
      await this.browser.saveAs(name, this.session.serialize('ods', name));
    } catch (e) {
      this.error(e);
    }
  }

  async importCsv(): Promise<void> {
    if (!(await this.confirmDiscard())) return;
    try {
      const opened = await this.browser.pick();
      if (opened) this.adopt(opened);
    } catch (e) {
      this.error(e);
    }
  }

  async saveToNas(name: string, bytes: Uint8Array): Promise<DocumentRef | null> {
    const where = await this.askNas('save', name);
    if (!where) return null;
    try {
      return await this.nas.saveNew(where.root, where.path, bytes);
    } catch (e) {
      if (!(e instanceof ConflictError) || e.reason !== 'exists') throw e;
      const choice = await this.askConflict(e);
      return choice === 'overwrite' ? this.nas.saveNew(where.root, where.path, bytes, true) : null;
    }
  }

  /**
   * Re-read the document from its store. `onlyIfClean` (external changes) gives up when edits
   * landed while the file was being read, instead of throwing them away.
   */
  async reloadFromStore(force = false, onlyIfClean = false): Promise<void> {
    const ref = this.session.document.peek();
    if (!ref) return;
    if (!force && !(await this.confirmDiscard())) return;
    const edits = this.session.editCount;
    try {
      const opened = await this.storeFor(ref).reload(ref);
      if (onlyIfClean && (this.session.editCount !== edits || this.session.document.peek() !== ref)) {
        this.info('The file changed on disk. Saving will ask before overwriting it.');
        return;
      }
      const kept = this.selection.state.peek();
      this.#keepSelection = true;
      try {
        this.session.reload(opened.bytes, opened.ref);
      } finally {
        this.#keepSelection = false;
      }
      this.selection.state.value = kept;
    } catch (e) {
      this.error(e);
    }
  }

  /** Electron noticed the file change on disk: reload silently when clean, ask when dirty. */
  async externalChange(): Promise<void> {
    if (!this.session.modified.peek() && !this.#saving) return this.reloadFromStore(true, true);
    this.info('The file changed on disk. Saving will ask before overwriting it.');
  }

  /** Before closing the window (Electron): save, discard, or stay. */
  async requestClose(): Promise<void> {
    if (this.#closing) return;
    this.#closing = true;
    try {
      if (await this.confirmDiscard()) {
        await clearRecovery();
        this.native?.confirmClose();
      } else this.native?.cancelClose();
    } finally {
      this.#closing = false;
    }
  }

  /** Record a save. `edits` is the session's edit count when the bytes were serialized. */
  saved(ref: DocumentRef, edits = this.session.editCount): void {
    this.session.markSaved(ref, edits);
    if (!this.session.modified.peek()) void clearRecovery();
    void this.remember(ref);
    this.native?.watch(ref.store === 'native' ? ref.id : null);
    this.info(`Saved ${ref.name}`);
  }

  async remember(ref: DocumentRef): Promise<void> {
    await rememberRecent(ref);
    await this.refreshRecents();
  }

  async refreshRecents(): Promise<void> {
    this.recents.value = await recentFiles();
  }

  async snapshot(): Promise<void> {
    // Serializing a big workbook takes a while: only redo it when something changed.
    const version = this.session.version.peek();
    if (!this.session.modified.peek() || version === this.#snapshotVersion) return;
    this.#snapshotVersion = version;
    try {
      await saveRecovery(
        this.session.title.peek(),
        this.session.document.peek(),
        this.session.serialize('ods'),
      );
    } catch {
      // Recovery is best effort.
    }
  }

  /** True when it is safe to replace the document (saved, discarded, or clean). */
  async confirmDiscard(): Promise<boolean> {
    if (!this.session.modified.peek()) return true;
    const choice = await new Promise<'save' | 'discard' | 'cancel'>(
      (resolve) => (this.dialog.value = { kind: 'confirm-discard', resolve }),
    );
    this.dialog.value = null;
    if (choice === 'cancel') return false;
    if (choice === 'save') return this.save();
    return true;
  }

  async askConflict(e: ConflictError): Promise<'overwrite' | 'saveAs' | 'reload' | 'cancel'> {
    const choice = await new Promise<'overwrite' | 'saveAs' | 'reload' | 'cancel'>(
      (resolve) => (this.dialog.value = { kind: 'conflict', reason: e.reason, message: e.message, resolve }),
    );
    this.dialog.value = null;
    return choice;
  }

  async askNas(
    mode: 'open' | 'save',
    suggestedName?: string,
  ): Promise<{ root: string; path: string } | null> {
    const r = await new Promise<{ root: string; path: string } | null>((resolve) => {
      this.dialog.value =
        suggestedName === undefined
          ? { kind: 'nas', mode, resolve }
          : { kind: 'nas', mode, suggestedName, resolve };
    });
    this.dialog.value = null;
    return r;
  }

  async prompt(title: string, label: string, value: string): Promise<string | null> {
    const r = await new Promise<string | null>(
      (resolve) => (this.dialog.value = { kind: 'prompt', title, label, value, resolve }),
    );
    this.dialog.value = null;
    this.grid.current?.focus();
    return r;
  }

  // ── Feedback ─────────────────────────────────────────────────────────────

  info(text: string): void {
    this.#showToast(text, 'info');
  }

  error(e: unknown): void {
    this.#showToast(message(e), 'error');
  }

  #showToast(text: string, tone: 'info' | 'error'): void {
    this.toast.value = { text, tone };
    if (this.#toastTimer) clearTimeout(this.#toastTimer);
    this.#toastTimer = setTimeout(() => (this.toast.value = null), tone === 'error' ? 6000 : 2500);
  }

  // ── Selection helpers ────────────────────────────────────────────────────

  get sheet(): Worksheet {
    return this.session.activeSheet.peek();
  }

  get sheetName(): string {
    return this.sheet.name;
  }

  get ranges(): CellRange[] {
    return this.selection.state.peek().ranges.map(toCellRange);
  }

  get range(): CellRange {
    return toCellRange(this.selection.range.peek());
  }

  get activeStyle(): CellStyle {
    const { row, column } = this.selection.active.peek();
    return this.session.workbook.peek().styles.get(this.sheet.effectiveStyleId(row, column));
  }

  /** The distinct styles in the selection (capped), for dialogs that show "(varies)". */
  selectionStyles(): CellStyle[] {
    const wb = this.session.workbook.peek();
    const sheet = this.sheet;
    const ids = new Set<number>();
    let seen = 0;
    for (const r of this.ranges) {
      const used = sheet.getUsedRange();
      const bottom = Math.min(r.bottom, Math.max(r.top, used?.bottom ?? r.top));
      const right = Math.min(r.right, Math.max(r.left, used?.right ?? r.left));
      for (let row = r.top; row <= bottom; row++)
        for (let c = r.left; c <= right; c++) {
          ids.add(sheet.effectiveStyleId(row, c));
          if (++seen > 20_000) return [...ids].map((i) => wb.styles.get(i));
        }
    }
    return [...ids].map((i) => wb.styles.get(i));
  }

  /** The selection, or — for a single cell — the data region around it (sort, dedupe). */
  get regionOrSelection(): CellRange {
    const r = this.range;
    if (!r.isSingleCell) return r;
    return DataRegion.currentRegion(this.sheet, r.top, r.left);
  }

  run(command: EditCommand | null): void {
    if (!command) return;
    try {
      this.session.execute(command);
    } catch (e) {
      this.error(e);
    }
  }

  // ── Cell edits ───────────────────────────────────────────────────────────

  commit(row: number, column: number, text: string): boolean {
    const s = this.session;
    const sheet = this.sheet;
    const rule = sheet.findValidation(row, column);
    if (rule) {
      const parsed = InputParser.parse(text, s.locale);
      if (parsed.formula === null && !Validator.isValid(rule, parsed.value, s.locale, this.resolver)) {
        if (rule.response === 'stop') {
          this.error(rule.errorMessage || 'That value is not allowed for this cell.');
          return false;
        }
        this.info(rule.errorMessage || 'That value does not match the cell’s validation rule.');
      }
    }
    const wb = s.workbook.peek();
    const set = EditCommands.setInput(this.sheetName, row, column, text, s.locale, wb);
    let expand: EditCommand | null = null;
    if (text !== '' && !sheet.findTable(row, column)) {
      for (const t of sheet.tables) {
        expand = DataEditCommands.expandTableForEdit(wb, this.sheetName, t, row, column);
        if (expand) break;
      }
    }
    this.run(expand ? new CompositeEditCommand(set.label, [set, expand]) : set);
    return true;
  }

  /** Reads a list rule's source range on the active sheet. */
  readonly resolver = (r: CellRange): Iterable<CellValue> => {
    const sheet = this.sheet;
    const version = this.session.version.peek();
    if (this.#listCache.version !== version || this.#listCache.sheet !== sheet)
      this.#listCache = { version, sheet, ranges: new Map() };
    const key = r.toString();
    const hit = this.#listCache.ranges.get(key);
    if (hit) return hit;
    // Preflight asks once per pasted cell, so read each source once per version, clipped to the
    // used range (a `Z:Z` source would otherwise be a million reads per cell).
    const out: CellValue[] = [];
    for (const [row, column] of clampToUsed(sheet, r)?.positions() ?? [])
      out.push(sheet.getValue(row, column));
    this.#listCache.ranges.set(key, out);
    return out;
  };
  #listCache: { version: number; sheet: Worksheet | null; ranges: Map<string, CellValue[]> } = {
    version: -1,
    sheet: null,
    ranges: new Map(),
  };

  /** Stop-rule check for paste/fill; true when the writes may go ahead. */
  #preflight(
    writes: Iterable<{ row: number; column: number; value: CellValue; isFormula: boolean }>,
  ): boolean {
    if (this.sheet.validations.length === 0) return true;
    const bad = ValidationPreflight.findStopViolations(
      this.sheet,
      writes,
      this.session.locale,
      this.resolver,
    );
    if (bad.length === 0) return true;
    const [r, c] = bad[0]!;
    const rule = this.sheet.findValidation(r, c);
    this.error(
      `${rule?.errorMessage || 'That value is not allowed for this cell.'} (${new CellAddress(c, r).toString()}${bad.length > 1 ? ` and ${bad.length - 1} more` : ''})`,
    );
    return false;
  }

  clearContents(): void {
    this.run(
      each(
        'Clear Contents',
        this.ranges.map((r) => EditCommands.clearContents(this.sheetName, r)),
      ),
    );
  }

  clearAll(): void {
    this.run(
      each(
        'Clear All',
        this.ranges.map((r) => EditCommands.clearAll(this.sheetName, r)),
      ),
    );
  }

  clearFormats(): void {
    const wb = this.session.workbook.peek();
    this.run(
      each(
        'Clear Formatting',
        this.ranges.map((r) => EditCommands.clearFormatting(wb, this.sheetName, r)),
      ),
    );
  }

  style(label: string, change: (s: CellStyle) => CellStyle): void {
    const wb = this.session.workbook.peek();
    this.run(
      each(
        label,
        this.ranges.map((r) => EditCommands.applyStyle(wb, this.sheetName, r, change, label)),
      ),
    );
  }

  /** Per-range border edits need the range's geometry (outline vs inner edges). */
  borders(label: string, edges: (r: CellRange) => [CellRange, (s: CellStyle) => CellStyle][]): void {
    const wb = this.session.workbook.peek();
    const commands = this.ranges.flatMap((r) =>
      edges(r).map(([sub, change]) => EditCommands.applyStyle(wb, this.sheetName, sub, change, label)),
    );
    this.run(each(label, commands));
  }

  edge(row: number, column: number, rowStep: number, columnStep: number): { row: number; column: number } {
    const [r, c] = EdgeNavigator.findEdge(this.sheet, row, column, rowStep, columnStep);
    return { row: r, column: c };
  }

  autoSum(): void {
    const { row, column } = this.selection.active.peek();
    const plans = AutoSumPlanner.plan(this.sheet, this.range, row, column);
    if (plans.length === 0) return this.info('Nothing to sum next to the selection.');
    const s = this.session;
    this.run(
      each(
        'AutoSum',
        plans.map((p) =>
          EditCommands.setInput(
            this.sheetName,
            p.targetRow,
            p.targetColumn,
            p.formula,
            s.locale,
            s.workbook.peek(),
          ),
        ),
      ),
    );
  }

  fill(direction: 'down' | 'right'): void {
    const r = this.range;
    if (direction === 'down' ? r.rowCount < 2 : r.columnCount < 2) return;
    const source =
      direction === 'down'
        ? CellRange.fromBounds(r.top, r.left, r.top, r.right)
        : CellRange.fromBounds(r.top, r.left, r.bottom, r.left);
    const fill = new FillCommand(this.sheetName, source, r);
    if (!this.#preflight(fill.projectWrites(this.session.workbook.peek()))) return;
    this.run(fill);
  }

  // ── Rows and columns ─────────────────────────────────────────────────────

  insertRows(): void {
    const r = this.range;
    this.run(new StructuralEditCommand(this.sheetName, 'row', 'insert', r.top, r.rowCount));
  }

  insertColumns(): void {
    const r = this.range;
    this.run(new StructuralEditCommand(this.sheetName, 'column', 'insert', r.left, r.columnCount));
  }

  deleteRows(): void {
    const r = this.range;
    this.run(
      new StructuralEditCommand(
        this.sheetName,
        'row',
        'delete',
        r.top,
        Math.min(r.rowCount, SheetDimensions.maxRows - r.top),
      ),
    );
  }

  deleteColumns(): void {
    const r = this.range;
    this.run(new StructuralEditCommand(this.sheetName, 'column', 'delete', r.left, r.columnCount));
  }

  hide(axis: 'rows' | 'columns', hidden: boolean): void {
    const r = this.range;
    if (axis === 'rows') {
      if (!hidden) return this.unhideRows();
      this.run(EditCommands.setRowsHidden(this.sheetName, r.top, r.bottom, true));
    } else {
      if (!hidden) return this.unhideColumns();
      this.run(EditCommands.setColumnsHidden(this.sheetName, r.left, r.right, true));
    }
  }

  unhideRows(): void {
    const runs = HiddenLines.rowsToUnhide(this.sheet, this.range.top, this.range.bottom);
    this.run(
      each(
        'Unhide Rows',
        runs.map(([a, b]) => EditCommands.setRowsHidden(this.sheetName, a, b, false)),
      ),
    );
  }

  unhideColumns(): void {
    const runs = HiddenLines.columnsToUnhide(this.sheet, this.range.left, this.range.right);
    this.run(
      each(
        'Unhide Columns',
        runs.map(([a, b]) => EditCommands.setColumnsHidden(this.sheetName, a, b, false)),
      ),
    );
  }

  async columnWidth(): Promise<void> {
    const r = this.range;
    const v = await this.prompt(
      'Column Width',
      'Width in pixels',
      String(Math.round(this.sheet.getColumnWidth(r.left))),
    );
    const n = Number(v);
    if (v === null || !Number.isFinite(n) || n <= 0) return;
    const cmds: EditCommand[] = [];
    for (let c = r.left; c <= Math.min(r.right, r.left + 1000); c++)
      cmds.push(EditCommands.setColumnWidth(this.sheetName, c, n));
    this.run(each('Column Width', cmds));
  }

  async rowHeight(): Promise<void> {
    const r = this.range;
    const v = await this.prompt(
      'Row Height',
      'Height in pixels',
      String(Math.round(this.sheet.getRowHeight(r.top))),
    );
    const n = Number(v);
    if (v === null || !Number.isFinite(n) || n <= 0) return;
    const cmds: EditCommand[] = [];
    for (let row = r.top; row <= Math.min(r.bottom, r.top + 10_000); row++)
      cmds.push(EditCommands.setRowHeight(this.sheetName, row, n));
    this.run(each('Row Height', cmds));
  }

  resizeColumn(column: number, width: number): void {
    const r = this.range;
    const whole =
      r.top === 0 && r.bottom >= SheetDimensions.maxRows - 1 && column >= r.left && column <= r.right;
    const columns = whole ? range(r.left, Math.min(r.right, r.left + 1000)) : [column];
    this.run(
      each(
        'Column Width',
        columns.map((c) => EditCommands.setColumnWidth(this.sheetName, c, width)),
      ),
    );
  }

  resizeRow(row: number, height: number): void {
    this.run(EditCommands.setRowHeight(this.sheetName, row, height));
  }

  autofitColumns(): void {
    const ctx = document.createElement('canvas').getContext('2d')!;
    const wb = this.session.workbook.peek();
    const sheet = this.sheet;
    const r = this.range;
    // One pass over the stored cells (not every row of every column), with the font and the
    // measured width cached: a column of repeated values measures each distinct text once.
    const fonts = new Map<number, string>();
    const measured = new Map<string, number>();
    const widths = new Map<number, number>();
    for (const [row, column, cell] of sheet.enumerateCells()) {
      if (column < r.left || column > r.right || row < r.top || row > r.bottom) continue;
      if (cell.value.kind === 'empty' || sheet.findMerge(row, column)) continue;
      const styleId = sheet.effectiveStyleId(row, column);
      const style = wb.styles.get(styleId);
      let font = fonts.get(styleId);
      if (font === undefined) {
        font = `${style.font.italic ? 'italic ' : ''}${style.font.bold ? 'bold ' : ''}${style.font.sizePoints * (4 / 3)}px ${style.font.family}, system-ui, sans-serif`;
        fonts.set(styleId, font);
      }
      const text = this.session.display(cell.value, style);
      const key = `${font}\u0000${text}`;
      let width = measured.get(key);
      if (width === undefined) {
        ctx.font = font;
        width = ctx.measureText(text).width;
        measured.set(key, width);
      }
      if (width > (widths.get(column) ?? 0)) widths.set(column, width);
    }
    const cmds = [...widths]
      .sort((a, b) => a[0] - b[0])
      .map(([c, width]) => EditCommands.setColumnWidth(this.sheetName, c, Math.ceil(width + 12)));
    this.run(each('Autofit Columns', cmds));
  }

  toggleFreeze(): void {
    const sheet = this.sheet;
    if (sheet.frozenRows > 0 || sheet.frozenColumns > 0)
      return this.run(EditCommands.setFrozenPanes(this.sheetName, 0, 0));
    const { row, column } = this.selection.active.peek();
    if (row === 0 && column === 0) return this.run(EditCommands.setFrozenPanes(this.sheetName, 1, 0));
    this.run(EditCommands.setFrozenPanes(this.sheetName, row, column));
  }

  merge(): void {
    const { row, column } = this.selection.active.peek();
    if (this.sheet.findMerge(row, column))
      return this.run(EditCommands.unmergeCells(this.sheetName, row, column));
    const r = this.range;
    if (r.isSingleCell) return;
    this.run(EditCommands.mergeCells(this.sheetName, r));
  }

  // ── Data ─────────────────────────────────────────────────────────────────

  sort(ascending: boolean): void {
    const region = this.regionOrSelection;
    if (region.rowCount < 2) return;
    const { column } = this.selection.active.peek();
    const hasHeader = this.range.isSingleCell && looksLikeHeader(this.sheet, region);
    this.run(
      DataEditCommands.sortByColumn(
        this.sheetName,
        region,
        column,
        ascending,
        hasHeader,
        this.session.locale,
      ),
    );
  }

  removeDuplicates(): void {
    const region = this.regionOrSelection;
    const hasHeader = looksLikeHeader(this.sheet, region);
    const columns = range(region.left, region.right);
    const { command, removedCount } = DataEditCommands.removeDuplicates(
      this.sheetName,
      region,
      columns,
      hasHeader,
    );
    this.run(command);
    const n = removedCount();
    this.info(n === 0 ? 'No duplicate rows found.' : `Removed ${n} duplicate row${n === 1 ? '' : 's'}.`);
  }

  async textToColumns(): Promise<void> {
    const r = this.range;
    const delimiter = await this.prompt('Text to Columns', 'Split on', ',');
    if (!delimiter) return;
    this.run(
      DataEditCommands.textToColumns(
        this.sheet,
        CellRange.fromBounds(r.top, r.left, r.bottom, r.left),
        delimiter.replace('\\t', '\t'),
      ),
    );
  }

  cleanUp(scope: 'sheet' | 'workbook'): void {
    const wb = this.session.workbook.peek();
    const plan =
      scope === 'sheet'
        ? CleanUp.forSheet(wb, this.sheet, this.session.locale)
        : CleanUp.forWorkbook(wb, this.session.locale);
    if (plan.isEmpty) return this.info('Nothing to clean up.');
    this.run(plan.toCommand(scope === 'sheet' ? 'Clean Up Sheet' : 'Clean Up Workbook'));
    this.info(plan.summary);
  }

  dataTarget(): DataTarget {
    const { row, column } = this.selection.active.peek();
    const t = this.sheet.findTable(row, column);
    if (t) return { range: t.range, hasHeader: t.hasHeader, table: t };
    const range = this.regionOrSelection;
    return { range, hasHeader: looksLikeHeader(this.sheet, range), table: null };
  }

  /** The AutoFilter or table at the cursor, if any (for enabling Clear Filter and friends). */
  get filterTarget(): DataTarget | null {
    const { row, column } = this.selection.active.peek();
    const sheet = this.sheet;
    const t = sheet.findTable(row, column);
    if (t) return { range: t.range, hasHeader: t.hasHeader, table: t };
    const auto = sheet.activeFilter;
    if (auto && auto.contains(row, column)) return { range: auto, hasHeader: true, table: null };
    return null;
  }

  sortFilterCommit(target: DataTarget, sort: SortSpec | null, filter: FilterSpec | null): void {
    this.run(this.#sortFilterCommand(target, sort, filter));
  }

  /**
   * One undo step. A plain AutoFilter left with no filter or sort would vanish, so it is put back in
   * the same step (keeping its chevrons).
   */
  #sortFilterCommand(target: DataTarget, sort: SortSpec | null, filter: FilterSpec | null): EditCommand {
    const sheet = this.sheetName;
    const locale = this.session.locale;
    const steps: EditCommand[] = [];
    if (target.table) {
      const auto = this.sheet.activeFilter;
      if (auto && auto.equals(target.table.range))
        steps.push(DataEditCommands.applySortAndFilter(sheet, auto, null, null, locale));
      steps.push(DataEditCommands.applyTableSortAndFilter(sheet, target.table, sort, filter, locale));
    } else {
      steps.push(DataEditCommands.applySortAndFilter(sheet, target.range, sort, filter, locale));
      if (!filter && !sort && this.sheet.activeFilter?.equals(target.range))
        steps.push(this.#autoFilterState(target.range));
    }
    return steps.length === 1 ? steps[0]! : new CompositeEditCommand('Sort & Filter', steps);
  }

  /** Filter dropdown: replace one column's checklist (null = show all) and optionally sort by it. */
  applyColumnFilter(
    target: DataTarget,
    column: number,
    allowed: ReadonlySet<string> | null,
    sortAscending?: boolean,
  ): void {
    const current = DataScope.effectiveFilter(this.sheet, target.range);
    const columns = new Map(current?.columns ?? []);
    const existing = columns.get(column);
    if (allowed) columns.set(column, { ...existing, allowedValues: allowed });
    else if (existing?.condition) columns.set(column, { condition: existing.condition });
    else columns.delete(column);
    const filter: FilterSpec | null = columns.size > 0 ? { columns, hasHeader: target.hasHeader } : null;
    const sort: SortSpec | null =
      sortAscending === undefined
        ? DataScope.effectiveSort(this.sheet, target.range)
        : { levels: [{ column, ascending: sortAscending }], hasHeader: target.hasHeader };
    this.sortFilterCommit(target, sort, filter);
  }

  /** The distinct values for a filter dropdown, honouring the other columns' filters. */
  filterValues(target: DataTarget, column: number): { key: string; label: string }[] {
    const spec = DataScope.effectiveFilter(this.sheet, target.range);
    const wb = this.session.workbook.peek();
    const values = FilterEngine.distinctValues(
      this.sheet,
      target.range,
      column,
      target.hasHeader,
      spec,
      this.session.locale,
    );
    const style = wb.styles.get(this.sheet.effectiveStyleId(target.range.top + 1, column));
    return values
      .map((v) => ({
        key: CellValue.toInvariantString(v),
        label: v.kind === 'empty' ? '(blank)' : this.session.display(v, style),
      }))
      .sort((a, b) =>
        a.key === '' ? 1 : b.key === '' ? -1 : a.label.localeCompare(b.label, this.session.locale),
      );
  }

  toggleFilter(): void {
    const sheet = this.sheet;
    if (sheet.activeFilter) {
      const range = sheet.activeFilter;
      this.run(
        each('Remove AutoFilter', [
          DataEditCommands.applySortAndFilter(this.sheetName, range, null, null, this.session.locale),
          this.#autoFilterState(null),
        ]),
      );
      return;
    }
    const target = this.dataTarget();
    if (target.table) return this.info('Tables have their own filter buttons.');
    // Whole columns (A:D) would filter, and save hidden flags for, a million empty rows.
    const range = clampToUsed(sheet, target.range);
    if (!range || range.rowCount < 2) return this.info('Select a range with a header row to filter.');
    this.#setAutoFilter(range);
  }

  clearFilter(): void {
    const target = this.filterTarget;
    if (!target) return;
    if (target.table) return this.sortFilterCommit(target, target.table.sort ?? null, null);
    this.sortFilterCommit(target, this.sheet.activeSort, null);
  }

  #setAutoFilter(range: CellRange): void {
    if (this.sheet.activeFilter?.equals(range)) return;
    this.run(this.#autoFilterState(range));
  }

  /** Undoable switch of the sheet's AutoFilter range (keeps spec/sort only when the range stays). */
  #autoFilterState(range: CellRange | null): EditCommand {
    const name = this.sheetName;
    let before: [CellRange | null, FilterSpec | null, SortSpec | null] = [null, null, null];
    const ws = (wb: { findSheet(n: string): Worksheet | undefined }) => wb.findSheet(name)!;
    return new ActionCommand(
      range ? 'AutoFilter' : 'Remove AutoFilter',
      (wb) => {
        const s = ws(wb);
        before = [s.activeFilter, s.activeFilterSpec, s.activeSort];
        const same = range && s.activeFilter?.equals(range);
        s.activeFilter = range;
        if (!same) {
          s.activeFilterSpec = null;
          s.activeSort = null;
        }
        return EditResult.none;
      },
      (wb) => {
        const s = ws(wb);
        [s.activeFilter, s.activeFilterSpec, s.activeSort] = before;
        return EditResult.none;
      },
    );
  }

  createTable(): void {
    const sheet = this.sheet;
    const r = clampToUsed(sheet, this.regionOrSelection);
    if (!r) return this.info('Select some data to make a table.');
    if (sheet.tables.some((t) => t.range.intersects(r)))
      return this.info('That range overlaps an existing table.');
    let n = 1;
    while (sheet.tables.some((t) => t.name.toLowerCase() === `table${n}`)) n++;
    const t: Table = { name: `Table${n}`, range: r, hasHeader: looksLikeHeader(sheet, r) };
    this.run(DataEditCommands.createStyledTable(this.session.workbook.peek(), this.sheetName, t));
    this.selection.select(r.top, r.left);
  }

  get activeTable(): Table | null {
    const { row, column } = this.selection.active.peek();
    return this.sheet.findTable(row, column) ?? null;
  }

  toggleHeaderColumn(): void {
    const t = this.activeTable;
    if (!t) return this.info('Put the cursor in a table first.');
    this.run(
      DataEditCommands.setTableHeaderColumn(
        this.session.workbook.peek(),
        this.sheetName,
        t,
        !t.hasHeaderColumn,
      ),
    );
  }

  toggleBandedRows(): void {
    const t = this.activeTable;
    if (!t) return this.info('Put the cursor in a table first.');
    this.run(DataEditCommands.setTableBandedRows(this.sheetName, t, !t.hasBandedRows));
  }

  removeTable(): void {
    const t = this.activeTable;
    if (!t) return;
    this.run(DataEditCommands.removeTable(this.sheetName, t));
  }

  // ── Validation, links, names ─────────────────────────────────────────────

  /** Replace the rules over the range with `rule` (null removes them). */
  setValidation(range: CellRange, rule: DataValidation | null): void {
    const name = this.sheetName;
    const old = this.sheet.validations.filter((v) => v.target.intersects(range));
    // Rules that reach outside the range keep those parts.
    const kept = old.flatMap((v) => subtractRange(v.target, range).map((target) => ({ ...v, target })));
    const remove = new ActionCommand(
      'Data Validation',
      (wb) => {
        const ws = wb.findSheet(name)!;
        for (const v of old) ws.removeValidation(v);
        for (const v of kept) ws.addValidation(v);
        return EditResult.none;
      },
      (wb) => {
        const ws = wb.findSheet(name)!;
        for (const v of kept) ws.removeValidation(v);
        for (const v of old) ws.addValidation(v);
        return EditResult.none;
      },
    );
    const steps: EditCommand[] = old.length > 0 ? [remove] : [];
    if (rule) steps.push(DataEditCommands.setValidation(name, rule));
    this.run(each('Data Validation', steps));
  }

  /** Items for a list rule's dropdown. */
  listItems(rule: DataValidation): string[] {
    if (rule.listValues && rule.listValues.length > 0) return [...rule.listValues];
    if (!rule.listSource) return [];
    const seen = new Set<string>();
    const out: string[] = [];
    for (const v of this.resolver(rule.listSource)) {
      if (v.kind === 'empty') continue;
      const t = v.kind === 'text' ? v.value : CellValue.toInvariantString(v);
      if (seen.has(t.toUpperCase())) continue;
      seen.add(t.toUpperCase());
      out.push(t);
    }
    return out;
  }

  setHyperlink(row: number, column: number, link: Hyperlink | null, text: string): void {
    const s = this.session;
    const wb = s.workbook.peek();
    const steps: EditCommand[] = [];
    const current = this.sheet.getValue(row, column);
    const shown = current.kind === 'empty' ? '' : s.editText(row, column);
    if (text !== shown && (text !== '' || link === null))
      steps.push(EditCommands.setInput(this.sheetName, row, column, text, s.locale, wb));
    else if (link && current.kind === 'empty')
      steps.push(EditCommands.setInput(this.sheetName, row, column, link.target, s.locale, wb));
    steps.push(EditCommands.setHyperlink(wb, this.sheetName, row, column, link));
    this.run(each('Set Hyperlink', steps));
  }

  /** ⌘-click: open the cell's link in a new tab/window. */
  openLink(row: number, column: number): boolean {
    const link = this.sheet.find(row, column)?.hyperlink;
    if (!link) return false;
    if (!/^(https?|mailto|file):/i.test(link.target)) {
      // An in-document reference like `#Sheet2.A1` or `Sheet2!A1`.
      if (this.goTo(link.target.replace(/^#/, '').replace(/\.(?=[A-Z]+\d)/i, '!'))) return true;
      this.error(`Can’t open “${link.target}”.`);
      return true;
    }
    // Electron's window-open handler sends web links to the system browser.
    window.open(link.target, '_blank', 'noopener');
    return true;
  }

  /** Named ranges change the workbook directly (not undoable), then everything recalculates. */
  addNamedRange(nr: NamedRange): string | null {
    const wb = this.session.workbook.peek();
    if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(nr.name) || CellAddress.tryParse(nr.name))
      return 'Names start with a letter or underscore and can’t look like a cell reference.';
    if (wb.findNamedRange(nr.name, nr.scope ?? null)) return `“${nr.name}” already exists.`;
    wb.addNamedRange(nr);
    this.session.recalculateAll();
    this.session.touch();
    return null;
  }

  removeNamedRange(nr: NamedRange): void {
    this.session.workbook.peek().removeNamedRange(nr.name, nr.scope ?? null);
    this.session.recalculateAll();
    this.session.touch();
  }

  /**
   * Clear every cell and reset the layout (line styles and hidden flags, filters, tables, validation,
   * merges, frozen panes) in one step. Asks first when that would discard more than 100 cells.
   */
  resetSheetData(): void {
    const sheet = this.sheet;
    const used = sheet.getUsedRange();
    if (!used) return this.info('The sheet is already empty.');
    if (sheet.cellCount > 100) {
      const ok = window.confirm(
        `Reset “${this.sheetName}”? This clears ${sheet.cellCount.toLocaleString()} cells and the sheet’s layout. You can undo it.`,
      );
      this.grid.current?.focus();
      if (!ok) return;
    }
    this.run(
      new CompositeEditCommand('Reset Sheet Data', [
        EditCommands.clearAll(this.sheetName, used),
        DataEditCommands.resetSheetLayout(this.sheetName),
      ]),
    );
  }

  insertChartSheet(chart: ChartDefinition | null): void {
    if (!chart) return this.info('Select a range with numbers to chart.');
    this.session.insertSheet((wb) => wb.addChartSheet(chart));
  }

  // ── Clipboard ────────────────────────────────────────────────────────────

  /** Copy to our block plus the system clipboard as TSV of displayed values. */
  copy(cut = false, data?: DataTransfer | null): void {
    const block = ClipboardBlock.capture(this.sheet, this.range, cut);
    const text = this.blockText(block);
    this.#clipboard = { block, text };
    if (data) data.setData('text/plain', text);
    else void navigator.clipboard?.writeText(text).catch(() => {});
    if (cut) this.info('Cut — paste to move.');
  }

  /** Paste: our own block when the system clipboard still holds what we copied, else TSV text. */
  paste(text: string | null, mode: 'all' | 'values' = 'all'): void {
    const own = this.#clipboard;
    const sheet = this.sheetName;
    if (own && (text === null || normalize(text) === normalize(own.text))) {
      const cut = own.block.isCut;
      const paste = new PasteBlockCommand(sheet, this.range, own.block, mode);
      if (!this.#preflight(paste.projectWrites())) return;
      this.run(paste);
      if (cut) this.#clipboard = null;
      return;
    }
    if (!text) return;
    const rows = parseCsv(text.replace(/\r\n?/g, '\n').replace(/\n$/, ''), {
      delimiter: text.includes('\t') ? '\t' : ',',
    });
    const locale = this.session.locale;
    const { row, column } = this.selection.active.peek();
    const wb = this.session.workbook.peek();
    const incoming = rows.flatMap((cells, r) =>
      cells.map((v, c) => {
        const parsed = InputParser.parse(v, locale);
        return { row: row + r, column: column + c, value: parsed.value, isFormula: parsed.formula !== null };
      }),
    );
    if (!this.#preflight(incoming)) return;
    const cmds = rows.flatMap((cells, r) =>
      cells.map((v, c) => EditCommands.setInput(sheet, row + r, column + c, v, locale, wb)),
    );
    this.run(each('Paste', cmds));
    const width = rows.reduce((w, r) => Math.max(w, r.length), 0);
    if (rows.length > 0 && width > 0)
      this.selection.selectRange({
        top: row,
        left: column,
        bottom: row + rows.length - 1,
        right: column + width - 1,
      });
  }

  async pasteFromMenu(mode: 'all' | 'values' = 'all'): Promise<void> {
    let text: string | null = null;
    try {
      text = (await navigator.clipboard?.readText()) ?? null;
    } catch {
      // Clipboard read denied: fall back to our own block.
    }
    this.paste(text, mode);
  }

  blockText(block: ClipboardBlock): string {
    const wb = this.session.workbook.peek();
    const sheet = this.sheet;
    return block.cells
      .map((row, r) =>
        row
          .map((cell, c) => {
            if (!cell) return '';
            const style = wb.styles.get(
              sheet.effectiveStyleId(block.sourceRange.top + r, block.sourceRange.left + c),
            );
            const t = this.session.display(cell.value, style);
            return /[\t\n"]/.test(t) ? `"${t.replaceAll('"', '""')}"` : t;
          })
          .join('\t'),
      )
      .join('\n');
  }

  armFormatPainter(): void {
    if (this.formatPainter.peek()) {
      this.formatPainter.value = null;
      return;
    }
    this.formatPainter.value = FormatBlock.capture(this.session.workbook.peek(), this.sheet, this.range);
    this.info('Select cells to paint the format onto.');
  }

  /** Called when a selection gesture ends: apply an armed format painter. */
  selectionDone(): void {
    const stamp = this.formatPainter.peek();
    if (!stamp) return;
    this.formatPainter.value = null;
    const r = this.range;
    const target = r.isSingleCell ? FormatBlock.paintTarget(r.top, r.left, r.top, r.left) : r;
    this.run(stamp.applyTo(this.sheetName, target));
  }

  // ── Find & go to ─────────────────────────────────────────────────────────

  /** Next cell (row-major, wrapping) whose shown text or formula contains the query. */
  findNext(query: string, backwards = false): boolean {
    if (!query) return false;
    const q = query.toLowerCase();
    const sheet = this.sheet;
    const wb = this.session.workbook.peek();
    // Walk the stored cells (row-major) rather than every position of the used range: one stray
    // cell at XFD1048576 would make that 17 billion lookups.
    const matches: [number, number][] = [];
    for (const [r, c, cell] of sheet.enumerateCells()) {
      if (cell.value.kind === 'empty' && cell.formula === null) continue;
      const shown = this.session.display(cell.value, wb.styles.get(sheet.effectiveStyleId(r, c)));
      if (shown.toLowerCase().includes(q) || (cell.formula ?? '').toLowerCase().includes(q))
        matches.push([r, c]);
    }
    if (matches.length === 0) return false;
    matches.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    const { row, column } = this.selection.active.peek();
    const after = (m: [number, number]) => m[0] > row || (m[0] === row && m[1] > column);
    const before = (m: [number, number]) => m[0] < row || (m[0] === row && m[1] < column);
    const hit = backwards
      ? ([...matches].reverse().find(before) ?? matches[matches.length - 1]!)
      : (matches.find(after) ?? matches[0]!);
    this.selection.select(hit[0], hit[1]);
    return true;
  }

  replaceAll(query: string, replacement: string): number {
    if (!query) return 0;
    const sheet = this.sheet;
    const s = this.session;
    const cmds: EditCommand[] = [];
    const writes: { row: number; column: number; value: CellValue; isFormula: boolean }[] = [];
    const pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    // Snapshot the stored cells first: the commands run later, but the walk must not see its own writes.
    for (const [r, c, cell] of [...sheet.enumerateCells()]) {
      const text =
        cell.formula !== null ? '=' + cell.formula : cell.value.kind === 'text' ? cell.value.value : null;
      if (text === null || !pattern.test(text)) continue;
      pattern.lastIndex = 0;
      const next = text.replace(pattern, () => replacement);
      const parsed = InputParser.parse(next, s.locale);
      writes.push({ row: r, column: c, value: parsed.value, isFormula: parsed.formula !== null });
      cmds.push(EditCommands.setInput(this.sheetName, r, c, next, s.locale, s.workbook.peek()));
    }
    if (!this.#preflight(writes)) return 0;
    this.run(each('Replace All', cmds));
    return cmds.length;
  }

  /** Name box: `B7`, `B2:D9`, `Sheet2!A1`. False when it doesn't parse. */
  goTo(reference: string): boolean {
    const text = reference.trim();
    const bang = text.lastIndexOf('!');
    if (bang > 0) {
      const name = text.slice(0, bang).replace(/^'|'$/g, '');
      const i = this.session.sheetNames.peek().findIndex((n) => n.toLowerCase() === name.toLowerCase());
      if (i < 0) return false;
      this.session.setActiveSheet(i);
    }
    const target = text.slice(bang + 1);
    const range =
      CellRange.tryParse(target) ??
      (CellAddress.tryParse(target) ? new CellRange(CellAddress.parse(target)) : null);
    if (!range) return false;
    batch(() =>
      this.selection.selectRange({
        top: range.top,
        left: range.left,
        bottom: range.bottom,
        right: range.right,
      }),
    );
    return true;
  }

  // ── Sheets ───────────────────────────────────────────────────────────────

  async renameSheet(index = this.session.activeIndex.peek()): Promise<void> {
    const current = this.session.sheetNames.peek()[index];
    if (current === undefined) return;
    const name = await this.prompt('Rename Sheet', 'Sheet name', current);
    if (name === null) return;
    const err = this.session.renameSheet(index, name.trim());
    if (err) this.error(err);
  }

  deleteSheet(index = this.session.activeIndex.peek()): void {
    if (this.session.sheetNames.peek().length <= 1) return this.info('A workbook needs at least one sheet.');
    this.session.deleteSheet(index);
  }
}

// ── helpers ──────────────────────────────────────────────────────────────

export function toCellRange(r: GridRange): CellRange {
  return CellRange.fromBounds(r.top, r.left, r.bottom, r.right);
}

function each(label: string, commands: EditCommand[]): EditCommand | null {
  if (commands.length === 0) return null;
  return commands.length === 1 ? commands[0]! : new CompositeEditCommand(label, commands);
}

function range(from: number, to: number): number[] {
  const out: number[] = [];
  for (let i = from; i <= to; i++) out.push(i);
  return out;
}

/** A first row of all text over a column that isn't all text reads as a header. */
function looksLikeHeader(sheet: Worksheet, region: CellRange): boolean {
  if (region.rowCount < 2) return false;
  let anyText = false;
  for (let c = region.left; c <= region.right; c++) {
    const v = sheet.getValue(region.top, c);
    if (v.kind !== 'text' && v.kind !== 'empty') return false;
    if (v.kind === 'text') anyText = true;
  }
  if (!anyText) return false;
  for (let c = region.left; c <= region.right; c++) {
    const below = sheet.getValue(region.top + 1, c);
    if (below.kind !== 'text' && below.kind !== 'empty') return true;
  }
  return false;
}

function splitNasId(id: string): [string, string] {
  const i = id.indexOf(':');
  return [id.slice(0, i), id.slice(i + 1)];
}

function normalize(text: string): string {
  return text.replace(/\r\n?/g, '\n').replace(/\n$/, '');
}

export function message(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}

/** The parts of `a` outside `b`, as up to four rectangles (above, below, left, right). */
function subtractRange(a: CellRange, b: CellRange): CellRange[] {
  if (!a.intersects(b)) return [a];
  const out: CellRange[] = [];
  const top = Math.max(a.top, b.top);
  const bottom = Math.min(a.bottom, b.bottom);
  if (a.top < b.top) out.push(CellRange.fromBounds(a.top, a.left, b.top - 1, a.right));
  if (a.bottom > b.bottom) out.push(CellRange.fromBounds(b.bottom + 1, a.left, a.bottom, a.right));
  if (a.left < b.left) out.push(CellRange.fromBounds(top, a.left, bottom, b.left - 1));
  if (a.right > b.right) out.push(CellRange.fromBounds(top, b.right + 1, bottom, a.right));
  return out;
}

/** The file is gone (not just unreadable right now): browser NotFoundError, NAS 404, native ENOENT. */
function isNotFound(e: unknown): boolean {
  if (!(e instanceof Error) && !(e instanceof DOMException)) return false;
  return e.name === 'NotFoundError' || /\bENOENT\b/.test(e.message);
}
