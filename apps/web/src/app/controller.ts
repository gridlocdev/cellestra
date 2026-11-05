import { batch, computed, effect, signal } from '@preact/signals';
import {
  AutoSumPlanner,
  CellRange,
  CellAddress,
  CleanUp,
  ClipboardBlock,
  CompositeEditCommand,
  DataEditCommands,
  DataRegion,
  EdgeNavigator,
  EditCommands,
  FillCommand,
  FormatBlock,
  HiddenLines,
  PasteBlockCommand,
  SheetDimensions,
  StructuralEditCommand,
  parseCsv,
  type CellStyle,
  type EditCommand,
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
  darkPalette,
  forgetRecent,
  lightPalette,
  loadRecovery,
  nativeBridge,
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
  | { readonly kind: 'shortcuts' }
  | { readonly kind: 'about' }
  | { readonly kind: 'settings' };

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

  constructor() {
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
    if (!this.native) {
      this.nasAvailable.value = (await this.nas.info()) !== null;
      addEventListener('beforeunload', (e) => {
        if (this.session.modified.peek()) e.preventDefault();
      });
    }
    void this.refreshRecents();
    setInterval(() => void this.snapshot(), RECOVERY_INTERVAL_MS);
    addEventListener('focus', () => this.session.recalculateVolatile());

    const snapshot = await loadRecovery().catch(() => null);
    if (snapshot) {
      const restore = await new Promise<boolean>(
        (resolve) => (this.dialog.value = { kind: 'recovery', name: snapshot.name, resolve }),
      );
      this.dialog.value = null;
      if (restore) {
        try {
          this.session.load(snapshot.bytes, snapshot.ref, snapshot.name.replace(/\.(csv|tsv)$/i, '.ods'));
          this.session.touch();
        } catch (e) {
          this.error(e);
        }
      } else await clearRecovery();
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

  async openFile(file: File): Promise<void> {
    if (!(await this.confirmDiscard())) return;
    try {
      this.adopt(await this.browser.openFile(file));
    } catch (e) {
      this.error(e);
    }
  }

  async openNativePath(path: string): Promise<void> {
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
      await forgetRecent(entry.ref);
      void this.refreshRecents();
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
  async save(force = false): Promise<boolean> {
    const ref = this.session.document.peek();
    if (!ref) return this.saveAs();
    const store = this.storeFor(ref);
    if (!store.capabilities.saveInPlace && !(ref.store === 'browser' && ref.handle)) return this.saveAs();
    const bytes = this.session.serialize();
    try {
      const saved = await store.save(ref, bytes, { force });
      this.saved(saved);
      return true;
    } catch (e) {
      if (!(e instanceof ConflictError)) {
        this.error(e);
        return false;
      }
      const choice = await this.askConflict(e);
      if (choice === 'overwrite') return this.save(true);
      if (choice === 'saveAs') return this.saveAs();
      if (choice === 'reload') await this.reloadFromStore(true);
      return false;
    }
  }

  async saveAs(format: 'ods' | 'csv' = this.session.format.peek()): Promise<boolean> {
    const base = this.session.title.peek().replace(/\.[^.]+$/, '');
    const name = `${base}.${format}`;
    const bytes = this.session.serialize(format, name);
    try {
      let saved: DocumentRef | null;
      const kind = this.primaryStore;
      if (kind === 'native') saved = await this.nativeStore!.saveAs(name, bytes);
      else if (kind === 'nas') saved = await this.saveToNas(name, bytes);
      else saved = await this.browser.saveAs(name, bytes);
      if (!saved) return false;
      // A CSV export of an .ods document is a copy: the document stays the .ods.
      if (format === 'csv' && this.session.format.peek() === 'ods') {
        this.info(`Exported ${saved.name}`);
        return true;
      }
      this.saved(saved);
      return true;
    } catch (e) {
      this.error(e);
      return false;
    }
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

  async reloadFromStore(force = false): Promise<void> {
    const ref = this.session.document.peek();
    if (!ref) return;
    if (!force && !(await this.confirmDiscard())) return;
    try {
      const opened = await this.storeFor(ref).reload(ref);
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
    if (!this.session.modified.peek()) return this.reloadFromStore(true);
    this.info('The file changed on disk. Saving will ask before overwriting it.');
  }

  /** Before closing the window (Electron): save, discard, or stay. */
  async requestClose(): Promise<void> {
    if (await this.confirmDiscard()) {
      void clearRecovery();
      this.native?.confirmClose();
    }
  }

  saved(ref: DocumentRef): void {
    this.session.markSaved(ref);
    void clearRecovery();
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
    if (!this.session.modified.peek()) return;
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

  commit(row: number, column: number, text: string): void {
    const s = this.session;
    this.run(EditCommands.setInput(this.sheetName, row, column, text, s.locale, s.workbook.peek()));
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
    this.run(new FillCommand(this.sheetName, source, r));
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
    const used = sheet.getUsedRange();
    if (!used) return;
    const r = this.range;
    const left = Math.max(r.left, used.left);
    const right = Math.min(r.right, used.right);
    const cmds: EditCommand[] = [];
    for (let c = left; c <= right; c++) {
      let width = 0;
      for (let row = used.top; row <= used.bottom; row++) {
        const cell = sheet.find(row, c);
        if (!cell || sheet.findMerge(row, c)) continue;
        const style = wb.styles.get(sheet.effectiveStyleId(row, c));
        ctx.font = `${style.font.italic ? 'italic ' : ''}${style.font.bold ? 'bold ' : ''}${style.font.sizePoints * (4 / 3)}px ${style.font.family}, system-ui, sans-serif`;
        width = Math.max(width, ctx.measureText(this.session.display(cell.value, style)).width);
      }
      if (width > 0) cmds.push(EditCommands.setColumnWidth(this.sheetName, c, Math.ceil(width + 12)));
    }
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
      this.run(new PasteBlockCommand(sheet, this.range, own.block, mode));
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
    const cmds = rows.flatMap((cells, r) =>
      cells.map((v, c) => EditCommands.setInput(sheet, row + r, column + c, v, locale, wb)),
    );
    this.run(each('Paste', cmds));
    const width = Math.max(0, ...rows.map((r) => r.length));
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
    const used = sheet.getUsedRange();
    if (!used) return false;
    const wb = this.session.workbook.peek();
    const width = used.right - used.left + 1;
    const total = (used.bottom - used.top + 1) * width;
    const { row, column } = this.selection.active.peek();
    const start =
      (Math.max(row, used.top) - used.top) * width +
      (Math.min(Math.max(column, used.left), used.right) - used.left);
    for (let i = 1; i <= total; i++) {
      const k = (((start + (backwards ? -i : i)) % total) + total) % total;
      const r = used.top + Math.floor(k / width);
      const c = used.left + (k % width);
      const cell = sheet.find(r, c);
      if (!cell) continue;
      const shown = this.session.display(cell.value, wb.styles.get(sheet.effectiveStyleId(r, c)));
      if (shown.toLowerCase().includes(q) || (cell.formula ?? '').toLowerCase().includes(q)) {
        this.selection.select(r, c);
        return true;
      }
    }
    return false;
  }

  replaceAll(query: string, replacement: string): number {
    if (!query) return 0;
    const sheet = this.sheet;
    const used = sheet.getUsedRange();
    if (!used) return 0;
    const s = this.session;
    const cmds: EditCommand[] = [];
    const pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    for (let r = used.top; r <= used.bottom; r++)
      for (let c = used.left; c <= used.right; c++) {
        const cell = sheet.find(r, c);
        if (!cell) continue;
        const text =
          cell.formula !== null ? '=' + cell.formula : cell.value.kind === 'text' ? cell.value.value : null;
        if (text === null || !pattern.test(text)) continue;
        pattern.lastIndex = 0;
        cmds.push(
          EditCommands.setInput(
            this.sheetName,
            r,
            c,
            text.replace(pattern, replacement),
            s.locale,
            s.workbook.peek(),
          ),
        );
      }
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
