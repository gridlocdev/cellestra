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
  DataRegion,
  listItems,
  namedRangeError,
  planCellEntry,
  SortFilterCommands,
  TableCommands,
  tableNameError,
  findTableByName,
  ChartEditCommands,
  ChartFields,
  type ChartField,
  DataScope,
  EdgeNavigator,
  FilterEngine,
  InputParser,
  ValidationPreflight,
  CellValue,
  EditCommands,
  FillCommand,
  FormatBlock,
  HiddenLines,
  MoveLinesCommand,
  PasteBlockCommand,
  sheetDeletionImpact,
  type SheetDeletionImpact,
  SheetDimensions,
  StructuralEditCommand,
  parseCsv,
  type Cell,
  type CellStyle,
  type DataValidation,
  type EditCommand,
  type FilterSpec,
  type Hyperlink,
  type NamedRange,
  type SortSpec,
  type Table,
  type Worksheet,
  ConditionalStyles,
  type Workbook,
  SetPageSetupCommand,
  printHtml,
  printableRange,
  type PageSetup,
  CsvOptions,
  ImportCsvCommand,
} from '@cellestra/core';
import type { ConflictError } from '@cellestra/ui';
import {
  BrowserStore,
  ChangePoller,
  NasStore,
  NativeStore,
  MAX_COLUMN,
  MAX_ROW,
  Selection,
  WorkbookSession,
  clearRecovery,
  csvLosses,
  darkPalette,
  forgetRecent,
  keepRecoveryAside,
  lightPalette,
  nativeBridge,
  canvasRowMeasure,
  contentRowHeights,
  printInFrame,
  orphanedRecoveries,
  recentFiles,
  rememberRecent,
  fitCsvColumns,
  fittedColumnWidths,
  saveRecovery,
  after,
  debounce,
  every,
  type CellestraNative,
  type DocumentRef,
  type GridController,
  type GridRange,
  type OpenedDocument,
  type RecentEntry,
} from '@cellestra/ui';

import type { ChartDefinition } from '@cellestra/charts';
import { DocumentStores } from './documentStores.ts';
import { DocumentLifecycle } from './documentLifecycle.ts';
import { message } from './lifecycleHost.ts';
import { findMatches, nextMatch, replaceIn, replaceableText, type FindMatch } from './find.ts';

/**
 * Built to be served by the NAS server (`VITE_CELLESTRA_NAS=1`). Other builds (desktop, a plain
 * static web build) never look for the server, so no NAS command, setting or dialog shows.
 */
export const NAS_BUILD = import.meta.env.VITE_CELLESTRA_NAS === '1';

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
  | {
      readonly kind: 'confirm';
      readonly title: string;
      readonly message: string;
      /** The label of the button that says yes. */
      readonly ok: string;
      readonly resolve: (ok: boolean) => void;
    }
  | { readonly kind: 'find' }
  | { readonly kind: 'recovery'; readonly name: string; readonly resolve: (restore: boolean) => void }
  | { readonly kind: 'outside-change'; readonly name: string; readonly resolve: (reload: boolean) => void }
  | {
      readonly kind: 'csv-loss';
      readonly name: string;
      readonly losses: readonly string[];
      readonly resolve: (choice: 'csv' | 'ods' | 'cancel') => void;
    }
  | { readonly kind: 'shortcuts' }
  | { readonly kind: 'conditionalFormats' }
  | { readonly kind: 'pageSetup' }
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

/** A chart field being picked on a data sheet: the chart sheet to return to and the field to set. */
export interface ChartPick {
  readonly chartSheet: Worksheet;
  readonly field: ChartField;
  /** How the prompt names the field: “the category labels”. */
  readonly label: string;
}

/** What the data commands act on: a table, else the current region (or selection). */
export interface DataTarget {
  readonly range: CellRange;
  readonly hasHeader: boolean;
  readonly table: Table | null;
}

const RECOVERY_INTERVAL_MS = 30_000;
/** A recovery snapshot is also taken this long after the last edit, so a pause in typing is safe. */
const RECOVERY_IDLE_MS = 3_000;
/** How much bigger than raw CSS px the grid draws at 100% zoom. */
const BASE_GRID_ZOOM = 1.3;
/** The zoom range the zoom commands allow, as fractions (25%–400%). */
export const MIN_ZOOM = 0.25;
export const MAX_ZOOM = 4;
const THEME_KEY = 'cellestra.theme';
const LAYOUT_KEY = 'cellestra.layout';
const RIBBON_COLLAPSED_KEY = 'cellestra.ribbonCollapsed';
const AUTO_SAVE_KEY = 'cellestra.autoSaveSeconds';

/** How fast pinching or ⌘/Ctrl-scrolling zooms: the zoom doubles every this many wheel px. */
const WHEEL_PX_PER_DOUBLING = 200;
/** A wheel event in lines (`deltaMode` 1) moves this many px per line. */
const WHEEL_LINE_PX = 16;

/** The zoom factor a pinch or ⌘/Ctrl-wheel event asks for: up (negative delta) zooms in. */
export function wheelZoomFactor(deltaY: number, deltaMode: number): number {
  const px = deltaMode === 1 ? deltaY * WHEEL_LINE_PX : deltaY;
  return 2 ** (-px / WHEEL_PX_PER_DOUBLING);
}

/** Auto-save's interval bounds, in seconds of no editing. 0 (or less) turns it off. */
export const AutoSaveSeconds = { default: 30, minimum: 5, maximum: 600 } as const;

/** The interval a stored or typed value means: 0 for off, otherwise clamped to 5–600. */
export function normalizeAutoSave(value: number): number {
  if (!Number.isFinite(value) || value <= 0) return 0;
  return Math.min(AutoSaveSeconds.maximum, Math.max(AutoSaveSeconds.minimum, Math.round(value)));
}

/**
 * What the Settings field means once edited: a number is normalized, while a cleared or unreadable
 * field keeps `current` (only a typed 0 turns auto-save off).
 */
export function autoSaveFromInput(text: string, current: number): number {
  const trimmed = text.trim();
  if (trimmed === '' || !Number.isFinite(Number(trimmed))) return current;
  return normalizeAutoSave(Number(trimmed));
}

/** The command chrome: a tabbed ribbon (the default) or the single compact toolbar. */
export type Layout = 'ribbon' | 'bar';

export { autoSaveFailure, message } from './lifecycleHost.ts';

/** PRD §8.7: clearing or deleting more non-empty cells than this asks first. */
export const LARGE_REMOVAL_CELLS = 100;

/** How many cells holding a value or formula lie in any of `ranges`, each counted once. */
export function filledCellCount(sheet: Worksheet, ranges: readonly CellRange[]): number {
  let count = 0;
  for (const [row, column, cell] of sheet.enumerateCells()) {
    if (cell.formula === null && cell.value.kind === 'empty') continue;
    if (ranges.some((r) => r.contains(row, column))) count++;
  }
  return count;
}

/** A stored preference; null when it is unset or storage is blocked (private window, policy). */
function stored(key: string): string | null {
  try {
    return localStorage.getItem(key);
  } catch {
    return null;
  }
}

/** Remember a preference; when storage is blocked it lasts for this session only. */
function store(key: string, value: string): void {
  try {
    localStorage.setItem(key, value);
  } catch {
    // Storage blocked: nothing to keep it in.
  }
}

/**
 * Everything the app does, independent of how it is drawn: documents and stores, edits on the
 * selection, clipboard, dialogs. Commands (`commands.ts`) and components call into this.
 */
export class AppController {
  readonly session = new WorkbookSession();
  readonly selection = new Selection();
  readonly grid: { current: GridController | null } = { current: null };
  /** The formula bar's Name Box: Go To (⌃G) puts the keyboard there. */
  readonly nameBox: { current: HTMLInputElement | null } = { current: null };

  readonly native: CellestraNative | null = nativeBridge();
  readonly stores = new DocumentStores({
    browser: new BrowserStore(),
    native: this.native ? new NativeStore(this.native) : null,
    nas: new NasStore(),
    nasAvailable: () => this.nasAvailable.peek(),
    nasDialogs: {
      choose: (mode, suggestedName) => this.askNas(mode, suggestedName),
      confirmOverwrite: async (e) => (await this.askConflict(e)) === 'overwrite',
    },
  });
  /** Whether the NAS server answered `/api/info` (the web client is being served by it). */
  readonly nasAvailable = signal(false);

  readonly dialog = signal<DialogState | null>(null);
  readonly toast = signal<{ text: string; tone: 'info' | 'error' } | null>(null);
  readonly paletteOpen = signal(false);
  readonly sheetPickerOpen = signal(false);
  readonly recentOpen = signal(false);
  /** Whether a dialog, menu or picker is open over the grid, which keeps the keyboard to itself. */
  modalOpen(): boolean {
    return !!(
      this.dialog.peek() ||
      this.paletteOpen.peek() ||
      this.sheetPickerOpen.peek() ||
      this.recentOpen.peek()
    );
  }
  readonly recents = signal<RecentEntry[]>([]);
  /** The zoom the user sees and sets: 1 is 100%. */
  readonly zoom = signal(1);
  /** The scale the grid draws at: 100% shows cells at a comfortable reading size, not raw CSS px. */
  readonly gridZoom = computed(() => this.zoom.value * BASE_GRID_ZOOM);
  readonly showFormulas = signal(false);
  readonly theme = signal<Theme>((stored(THEME_KEY) as Theme | null) ?? 'system');
  readonly prefersDark = signal(matchMedia('(prefers-color-scheme: dark)').matches);
  readonly isDark = computed(() =>
    this.theme.value === 'system' ? this.prefersDark.value : this.theme.value === 'dark',
  );
  readonly palette = computed(() => (this.isDark.value ? darkPalette : lightPalette));
  readonly layout = signal<Layout>(stored(LAYOUT_KEY) === 'bar' ? 'bar' : 'ribbon');
  /** The ribbon shows only its tabs; clicking one opens it. */
  readonly ribbonCollapsed = signal(stored(RIBBON_COLLAPSED_KEY) === 'true');
  /** Seconds of no editing before the document is saved in place; 0 is off. */
  readonly autoSaveSeconds = signal(
    normalizeAutoSave(Number(stored(AUTO_SAVE_KEY) ?? AutoSaveSeconds.default)),
  );
  readonly formatPainter = signal<FormatBlock | null>(null);
  /** A chart field being picked on the sheet (the chart panel's "Pick on sheet"), or null. */
  readonly chartPick = signal<ChartPick | null>(null);
  readonly findQuery = signal('');
  /** The copy/cut source the grid outlines with marching ants, until paste, Escape, or an edit. */
  readonly marquee = signal<{ sheet: string; range: GridRange; cut: boolean } | null>(null);
  /** A style drawn over the selection while a font size or colour is hovered; nothing is applied. */
  readonly stylePreview = signal<((s: CellStyle) => CellStyle) | null>(null);
  /** The grid's right-click menu: what was clicked and where. */
  readonly contextMenu = signal<{ zone: 'cell' | 'columnHeader' | 'rowHeader'; x: number; y: number } | null>(
    null,
  );

  /** Our last copy or cut, stamped with the document generation it came from. */
  #clipboard: { block: ClipboardBlock; text: string; generation: number } | null = null;
  /** Set while reloading the same document, so the cursor stays put. */
  #keepSelection = false;
  #cancelToast: (() => void) | null = null;
  /** The session version the last recovery snapshot was taken at. */
  #snapshotVersion = -1;
  /** Opening, saving, auto-save, reloading and closing the document. */
  readonly lifecycle: DocumentLifecycle;
  /** Settles when `start()` has finished (recovery offered): OS file opens wait for it. */
  readonly #ready: Promise<void>;
  #markReady: () => void = () => {};

  constructor() {
    this.#ready = new Promise<void>((resolve) => (this.#markReady = resolve));
    this.session.captureSelection = () => this.selection.state.peek();
    this.session.restoreSelection = (s) => (this.selection.state.value = s);
    this.session.onRefused = (text) => this.#showToast(text, 'error');
    // Another dialog replacing an unanswered question answers it with its cancel value.
    effect(() => {
      const shown = this.dialog.value;
      if (this.#asking && shown !== this.#asking.dialog) this.#asking.cancel();
    });
    matchMedia('(prefers-color-scheme: dark)').addEventListener(
      'change',
      (e) => (this.prefersDark.value = e.matches),
    );
    effect(() => {
      store(THEME_KEY, this.theme.value);
      document.documentElement.dataset.theme = this.isDark.value ? 'dark' : 'light';
    });
    effect(() => {
      store(LAYOUT_KEY, this.layout.value);
      store(RIBBON_COLLAPSED_KEY, String(this.ribbonCollapsed.value));
    });
    effect(() => store(AUTO_SAVE_KEY, String(this.autoSaveSeconds.value)));
    this.lifecycle = new DocumentLifecycle({
      session: this.session,
      stores: this.stores,
      autoSaveSeconds: this.autoSaveSeconds,
      native: () => this.native,
      questions: {
        discard: () =>
          this.ask<'save' | 'discard' | 'cancel'>('cancel', (resolve) => ({
            kind: 'confirm-discard',
            resolve,
          })),
        csvLoss: (name) => this.askCsvLoss(name),
        conflict: (e) => this.askConflict(e),
      },
      commitOpenEdit: () => this.commitOpenEdit(),
      isEditing: () => this.grid.current?.isEditing() ?? false,
      dialogOpen: () => this.dialog.peek() !== null,
      info: (text) => this.info(text),
      error: (e) => this.error(e),
      recovery: { clear: () => clearRecovery(), keepAside: keepRecoveryAside },
      remember: (ref) => this.remember(ref),
      replaced: () => {
        this.clearMarquee();
        this.grid.current?.focus();
      },
      reloadInPlace: (run) => {
        const kept = this.selection.state.peek();
        this.#keepSelection = true;
        try {
          run();
        } finally {
          this.#keepSelection = false;
        }
        this.selection.state.value = kept;
      },
    });
    // A pick ends when the user leaves for a chart sheet or another workbook opens.
    effect(() => {
      const sheet = this.session.activeSheet.value;
      const pick = this.chartPick.peek();
      if (!pick) return;
      if (sheet.isChartSheet || this.session.workbook.value.indexOf(pick.chartSheet) < 0)
        this.chartPick.value = null;
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
    // A file opens at the zoom it was saved with (100% when it has none)…
    effect(() => {
      const saved = this.session.workbook.value.settings.zoom;
      this.zoom.value = saved === null ? 1 : Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, saved / 100));
    });
    // …and saves the one in use. Zooming isn't an edit: it doesn't mark the document modified.
    effect(() => {
      const percent = Math.round(this.zoom.value * 100);
      const settings = this.session.workbook.peek().settings;
      if (settings.zoom !== null || percent !== 100) settings.zoom = percent;
    });
  }

  /** Detect the NAS server (NAS builds only), wire Electron, offer crash recovery, open a file passed by the OS. */
  async start(): Promise<void> {
    try {
      if (!this.native) {
        if (NAS_BUILD) this.nasAvailable.value = (await this.nas.info()) !== null;
        if (this.nasAvailable.peek()) this.#pollNasDocument();
        addEventListener('beforeunload', (e) => {
          if (this.hasUnsavedWork) e.preventDefault();
        });
      }
      void this.refreshRecents();
      every(RECOVERY_INTERVAL_MS, () => void this.snapshot());
      // Snapshot once editing pauses, and when the tab is hidden or closed (the last moment a
      // write to IndexedDB still reliably lands), so a crash loses seconds, not half a minute.
      const idle = debounce(RECOVERY_IDLE_MS, () => void this.snapshot());
      effect(() => {
        void this.session.version.value;
        idle.call();
      });
      document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'hidden') void this.snapshot();
      });
      addEventListener('pagehide', () => void this.snapshot());
      addEventListener('focus', () => {
        this.session.recalculateVolatile();
        this.focusGridIfIdle();
      });
      await this.offerRecovery();
    } finally {
      this.#markReady();
      // Typing right after launch edits the active cell, with no click first.
      queueMicrotask(() => this.focusGridIfIdle());
    }
  }

  /** Give the grid the keyboard when nothing else has it (focus sits on the page itself). */
  focusGridIfIdle(): void {
    if (typeof document === 'undefined' || this.dialog.peek()) return;
    const active = document.activeElement;
    if (!active || active === document.body || active === document.documentElement)
      this.grid.current?.focus();
  }

  /**
   * Offer the snapshots no open tab owns (each tab keeps its own), newest first. Restoring one
   * stops the round; the rest wait for the next start. Discarded ones are deleted.
   */
  async offerRecovery(): Promise<void> {
    const orphans = await orphanedRecoveries().catch(() => []);
    for (const snapshot of orphans) {
      const restore = await this.ask(false, (resolve) => ({
        kind: 'recovery',
        name: snapshot.name,
        resolve,
      }));
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

  get browser(): BrowserStore {
    return this.stores.browser;
  }

  get nas(): NasStore {
    return this.stores.nas;
  }

  newDocument(): Promise<void> {
    return this.lifecycle.newDocument();
  }

  open(): Promise<void> {
    return this.lifecycle.open(() => this.stores.primary.pick());
  }

  /** Upload from this computer even when the NAS is the primary store. */
  openLocal(): Promise<void> {
    return this.lifecycle.open(() => this.browser.pick());
  }

  /** Open a dropped file; with a File System Access handle it can be saved back in place. */
  openFile(file: File, handle?: Promise<FileSystemHandle | null>): Promise<void> {
    return this.lifecycle.open(async () => {
      const h = await handle?.catch(() => null);
      return this.browser.openFile(file, h?.kind === 'file' ? (h as FileSystemFileHandle) : undefined);
    });
  }

  async openNativePath(path: string): Promise<void> {
    // A file passed at launch waits for the recovery prompt, so neither replaces the other unasked.
    await this.#ready;
    const native = this.stores.native;
    if (!native) return;
    return this.lifecycle.open(() => native.openPath(path));
  }

  openRecent(entry: RecentEntry): Promise<void> {
    return this.lifecycle.open(
      () => this.stores.reopen(entry.ref),
      async () => {
        await forgetRecent(entry);
        void this.refreshRecents();
      },
    );
  }

  adopt(opened: OpenedDocument): void {
    this.lifecycle.adopt(opened);
  }

  /** Save in place when the store can; otherwise (new document, browser without a handle) Save As. */
  save(force = false): Promise<boolean> {
    return this.lifecycle.save(force);
  }

  /** The auto-save tick (see `DocumentLifecycle.autoSave`). Returns whether it wrote. */
  autoSave(): Promise<boolean> {
    return this.lifecycle.autoSave();
  }

  saveAs(format: 'ods' | 'csv' = this.session.format.peek()): Promise<boolean> {
    return this.lifecycle.saveAs(format);
  }

  /**
   * Saving a CSV document keeps one sheet's values. When that loses something (other sheets,
   * formulas, formatting), ask first. Resolves 'csv' straight away when nothing would be lost.
   */
  async askCsvLoss(name: string): Promise<'csv' | 'ods' | 'cancel'> {
    const losses = csvLosses(this.session.workbook.peek(), this.session.csvSheet());
    if (losses.length === 0) return 'csv';
    return this.ask<'csv' | 'ods' | 'cancel'>('cancel', (resolve) => ({
      kind: 'csv-loss',
      name,
      losses,
      resolve,
    }));
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

  /**
   * Import CSV (PRD §8.4): pick a .csv or .tsv and put its data in the active sheet, replacing what
   * is there, in one undoable step. The open document stays the open document. Asks first when
   * more than 100 filled cells would go.
   */
  async importCsv(): Promise<void> {
    if (!this.commitOpenEdit()) return;
    let picked: OpenedDocument | null;
    try {
      picked = await this.browser.pick();
    } catch (e) {
      return this.error(e);
    }
    if (!picked) return;
    const name = picked.ref.name;
    if (!/\.(csv|tsv|txt)$/i.test(name))
      return this.error(`Import CSV takes a .csv or .tsv file, not ${name}.`);
    const sheet = this.sheet;
    if (sheet.cellCount > LARGE_REMOVAL_CELLS) {
      const message = `This replaces ${sheet.cellCount.toLocaleString()} cells on “${this.sheetName}” with the data in ${name}. You can undo it.`;
      if (!(await this.confirm('Import CSV?', message, 'Import'))) return;
    }
    const text = new TextDecoder().decode(picked.bytes);
    const ctx = this.session.textMeasure();
    this.run(
      new ImportCsvCommand(this.sheetName, text, CsvOptions.forFileName(name), (wb, sheet) => {
        if (ctx) fitCsvColumns(ctx, wb, sheet);
      }),
    );
  }

  /** Re-read the document from its store; `onlyIfClean` gives up when edits landed meanwhile. */
  reloadFromStore(force = false, onlyIfClean = false): Promise<void> {
    return this.lifecycle.reloadFromStore(force, onlyIfClean);
  }

  /** Electron noticed the file change on disk (see `DocumentLifecycle.externalChange`). */
  externalChange(generation = this.session.generation): Promise<void> {
    return this.lifecycle.externalChange(generation);
  }

  /**
   * Poll the open NAS document's ETag, so a save from somewhere else is noticed (the server can't
   * push one). Follows whichever document is open; a hidden tab doesn't poll.
   */
  #pollNasDocument(): void {
    let poller: ChangePoller | null = null;
    let watching: string | null = null;
    effect(() => {
      const ref = this.session.document.value;
      const id = ref?.store === 'nas' ? ref.id : null;
      if (id === watching) return;
      poller?.stop();
      poller = null;
      watching = id;
      if (id === null) return;
      poller = new ChangePoller({
        fetchVersion: () => this.nas.currentVersion(id),
        knownVersion: () => {
          const open = this.session.document.peek();
          return open?.id === id && open.store === 'nas' ? open.version : null;
        },
        onChange: () => this.nasFileChanged(),
        isVisible: () => document.visibilityState === 'visible',
      });
      poller.start();
    });
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible') poller?.wake();
    });
  }

  /**
   * The open NAS file was saved by someone else. Without local edits, offer to reload it; with
   * them, say so: the next save meets the server's conflict check and asks then. Returns false
   * when now is a bad moment (a save running, a dialog or a cell edit open), so the poller
   * reports the change again on its next check.
   */
  nasFileChanged(): boolean {
    if (this.busy) return false;
    const name = this.session.title.peek();
    if (this.session.modified.peek()) {
      this.#showToast(`“${name}” changed on the server. Saving will ask before overwriting it.`, 'error');
      return true;
    }
    void this.ask(false, (resolve) => ({ kind: 'outside-change', name, resolve })).then((reload) => {
      if (reload) return this.reloadFromStore(true, true);
    });
    return true;
  }

  /** Before closing the window (Electron): save, discard, or stay. */
  requestClose(): Promise<void> {
    return this.lifecycle.requestClose();
  }

  /** Close the document (⌘W): the window on desktop, back to a blank workbook in a browser. */
  closeDocument(): Promise<void> {
    return this.lifecycle.closeDocument();
  }

  async remember(ref: DocumentRef): Promise<void> {
    await rememberRecent(ref);
    await this.refreshRecents();
  }

  async refreshRecents(): Promise<void> {
    // A NAS file listed on this origin by an earlier NAS build can't be opened without the server.
    const nas = this.nasAvailable.peek();
    this.recents.value = (await recentFiles()).filter((r) => nas || r.ref.store !== 'nas');
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

  /** The question dialog waiting for an answer, and how to give it its cancel answer. */
  #asking: { readonly dialog: DialogState; readonly cancel: () => void } | null = null;

  /** Whether a question dialog (save changes?, conflict, name…) is waiting for an answer. */
  get asking(): boolean {
    return this.#asking !== null && this.dialog.peek() === this.#asking.dialog;
  }

  /**
   * Show a question dialog and wait for its answer. A question asked while another is open
   * answers the open one with `cancel` first, so whatever awaited it (a save) never hangs.
   */
  ask<T>(cancel: T, make: (resolve: (value: T) => void) => DialogState): Promise<T> {
    this.#asking?.cancel();
    return new Promise<T>((resolve) => {
      let settled = false;
      const settle = (value: T) => {
        if (settled) return;
        settled = true;
        if (this.#asking?.dialog === dialog) this.#asking = null;
        if (this.dialog.peek() === dialog) this.dialog.value = null;
        resolve(value);
      };
      const dialog = make(settle);
      this.#asking = { dialog, cancel: () => settle(cancel) };
      this.dialog.value = dialog;
    });
  }

  /** Ask a yes-or-no question in the app's own dialog. Closing it answers no; either way the grid gets focus back. */
  async confirm(title: string, message: string, ok: string): Promise<boolean> {
    const answer = await this.ask(false, (resolve) => ({ kind: 'confirm', title, message, ok, resolve }));
    this.grid.current?.focus();
    return answer;
  }

  /**
   * Commit the cell edit that is open, if any, so what was typed counts as a change before the
   * document is saved, replaced or closed. False when the value was refused (the editor stays open).
   */
  commitOpenEdit(): boolean {
    return this.grid.current?.commitEdit() ?? true;
  }

  /** A bad moment to replace the document unasked: a save, a dialog or a cell edit is open. */
  get busy(): boolean {
    return this.lifecycle.busy;
  }

  /** Changes that closing the tab would lose: unsaved edits, or a cell edit not committed yet. */
  get hasUnsavedWork(): boolean {
    return this.lifecycle.hasUnsavedWork;
  }

  /** True when it is safe to replace the document (saved, discarded, or clean). */
  confirmDiscard(): Promise<boolean> {
    return this.lifecycle.confirmDiscard();
  }

  async askConflict(e: ConflictError): Promise<'overwrite' | 'saveAs' | 'reload' | 'cancel'> {
    return this.ask<'overwrite' | 'saveAs' | 'reload' | 'cancel'>('cancel', (resolve) => ({
      kind: 'conflict',
      reason: e.reason,
      message: e.message,
      resolve,
    }));
  }

  async askNas(
    mode: 'open' | 'save',
    suggestedName?: string,
  ): Promise<{ root: string; path: string } | null> {
    return this.ask<{ root: string; path: string } | null>(null, (resolve) =>
      suggestedName === undefined
        ? { kind: 'nas', mode, resolve }
        : { kind: 'nas', mode, suggestedName, resolve },
    );
  }

  async prompt(title: string, label: string, value: string): Promise<string | null> {
    const r = await this.ask<string | null>(null, (resolve) => ({
      kind: 'prompt',
      title,
      label,
      value,
      resolve,
    }));
    this.grid.current?.focus();
    return r;
  }

  /**
   * Scale the zoom (a pinch, ⌘-wheel), kept within the zoom range. Not rounded: a trackpad sends
   * many tiny steps that rounding would swallow.
   */
  zoomBy(factor: number): void {
    this.zoom.value = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, this.zoom.peek() * factor));
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
    this.#cancelToast?.();
    this.#cancelToast = after(tone === 'error' ? 6000 : 2500, () => (this.toast.value = null));
  }

  // ── Selection helpers ────────────────────────────────────────────────────

  get sheet(): Worksheet {
    return this.session.activeSheet.peek();
  }

  #conditional: { workbook: Workbook; version: number; styles: ConditionalStyles } | null = null;

  /** The style a cell shows with its conditional formats applied (made fresh after each change). */
  conditionalStyle(style: CellStyle, row: number, column: number): CellStyle {
    const workbook = this.session.workbook.peek();
    const version = this.session.version.peek();
    const c = this.#conditional;
    if (!c || c.workbook !== workbook || c.version !== version)
      this.#conditional = { workbook, version, styles: new ConditionalStyles(workbook) };
    return this.#conditional!.styles.resolve(style, this.sheetName, row, column);
  }

  // ── Printing ─────────────────────────────────────────────────────────────

  setPageSetup(setup: PageSetup): void {
    this.run(new SetPageSetupCommand(this.sheetName, setup));
  }

  /** The active sheet (or just the selection) laid out in pages, or null for a chart sheet. */
  #printDocument(selectionOnly: boolean, setup: PageSetup | null = this.sheet.pageSetup) {
    const sheet = this.sheet;
    if (sheet.isChartSheet) {
      this.info('Chart sheets can’t be printed yet. Print a sheet with cells instead.');
      return null;
    }
    const r = this.range;
    const range = selectionOnly
      ? printableRange(sheet, { top: r.top, left: r.left, bottom: r.bottom, right: r.right })
      : null;
    if (selectionOnly && !range) {
      this.info('The selection is empty: there is nothing to print.');
      return null;
    }
    const styles = this.session.workbook.peek().styles;
    const display = (value: CellValue, style: CellStyle) => this.session.display(value, style);
    // Rows the grid shows taller for their content (wrapped text, big fonts) print that tall too.
    const ctx = document.createElement('canvas').getContext('2d');
    return printHtml({
      sheet,
      styles,
      setup,
      range,
      display,
      contentHeights: contentRowHeights(sheet, styles, canvasRowMeasure(sheet, display, ctx)),
      conditional: (style, row, column) => this.conditionalStyle(style, row, column),
      title: `${this.session.title.peek().replace(/\.[^.]+$/, '')} — ${sheet.name}`,
    });
  }

  /** Print the active sheet, or the selection, through the system's (or browser's) print dialog. */
  async print(selectionOnly = false, setup?: PageSetup): Promise<void> {
    const doc = this.#printDocument(selectionOnly, setup);
    if (!doc) return;
    try {
      if (this.native?.print) {
        await this.native.print(doc.html, { width: doc.setup.paperWidth, height: doc.setup.paperHeight });
      } else await printInFrame(doc.html);
    } catch (e) {
      this.error(e);
    }
  }

  /** Save the active sheet as a PDF (in a browser: the print dialog, where Save as PDF is). */
  async exportPdf(): Promise<void> {
    const doc = this.#printDocument(false);
    if (!doc) return;
    try {
      if (this.native?.exportPdf) {
        const base = this.session.title.peek().replace(/\.[^.]+$/, '');
        const saved = await this.native.exportPdf(doc.html, `${base}.pdf`);
        if (saved) this.info(`Exported ${saved.split(/[\\/]/).pop()}`);
        return;
      }
      this.info('Choose “Save as PDF” as the printer to save a PDF.');
      await printInFrame(doc.html);
    } catch (e) {
      this.error(e);
    }
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

  run(command: EditCommand | null, target?: GridRange | readonly GridRange[]): void {
    if (!command) return;
    try {
      this.session.execute(command, target);
    } catch (e) {
      this.error(e);
    }
    // Any edit but a paste ends the copy marquee (and with it a pending cut).
    if (!(command instanceof PasteBlockCommand)) this.clearMarquee();
  }

  /** Escape or an edit: stop outlining the copied range; a pending cut is dropped. */
  clearMarquee(): void {
    if (!this.marquee.peek()) return;
    this.marquee.value = null;
    if (this.#clipboard?.block.isCut) this.#clipboard = null;
  }

  /** Hover preview of a style change over the selection; null puts the real styles back. */
  previewStyle(change: ((s: CellStyle) => CellStyle) | null): void {
    this.stylePreview.value = change;
  }

  // ── Cell edits ───────────────────────────────────────────────────────────

  commit(row: number, column: number, text: string): boolean {
    const s = this.session;
    const entry = planCellEntry(
      s.workbook.peek(),
      this.sheetName,
      row,
      column,
      text,
      s.locale,
      this.resolver,
    );
    const rule = entry.brokenRule;
    if (rule?.response === 'stop') {
      this.error(rule.errorMessage || 'That value is not allowed for this cell.');
      return false;
    }
    if (rule && rule.response !== 'none')
      this.info(rule.errorMessage || 'That value does not match the cell’s validation rule.');
    this.run(entry.command);
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

  /**
   * PRD §8.7: removing data from more than 100 non-empty cells asks first; fewer just happen
   * (undo covers them). Null when the removal may go ahead at once; otherwise the question to ask.
   */
  #largeRemoval(ranges: readonly CellRange[], what: string, ok: string): Promise<boolean> | null {
    const count = filledCellCount(this.sheet, ranges);
    if (count <= LARGE_REMOVAL_CELLS) return null;
    return this.confirm(
      `${what}?`,
      `This removes data from ${count.toLocaleString()} cells. You can undo it.`,
      ok,
    );
  }

  async clearContents(): Promise<void> {
    const asking = this.#largeRemoval(this.ranges, 'Clear the contents', 'Clear');
    if (asking && !(await asking)) return;
    this.run(
      each(
        'Clear Contents',
        this.ranges.map((r) => EditCommands.clearContents(this.sheetName, r)),
      ),
    );
  }

  async clearAll(): Promise<void> {
    const asking = this.#largeRemoval(this.ranges, 'Clear everything in the selection', 'Clear');
    if (asking && !(await asking)) return;
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
    this.stylePreview.value = null;
    // A chart sheet shows no cells: a style would land on an invisible A1.
    if (this.sheet.isChartSheet) return;
    const wb = this.session.workbook.peek();
    this.run(
      each(
        label,
        this.ranges.map((r) => EditCommands.applyStyle(wb, this.sheetName, r, change, label)),
      ),
    );
  }

  /** Turn the selection into checkboxes, or back into plain cells when the active cell is one. */
  toggleCheckboxStyle(): void {
    const wb = this.session.workbook.peek();
    const remove = !!this.activeStyle.checkbox;
    this.run(
      each(
        remove ? 'Remove Checkbox' : 'Insert Checkbox',
        this.ranges.map((r) =>
          remove
            ? EditCommands.removeCheckboxes(wb, this.sheetName, r)
            : EditCommands.insertCheckboxes(wb, this.sheetName, r),
        ),
      ),
    );
  }

  /**
   * Space, or a click on a box: tick the checkboxes at `at` (or in the selection) when any of them
   * is clear, otherwise clear them all. Only TRUE/FALSE cells count, and a formula's result can't
   * be ticked. False when there is no checkbox to toggle, so Space can do its usual thing.
   */
  toggleCheckboxes(at?: { row: number; column: number }): boolean {
    const wb = this.session.workbook.peek();
    const sheet = this.sheet;
    const positions: [number, number][] = [];
    let anyClear = false;
    const visit = (row: number, column: number) => {
      const cell = sheet.find(row, column);
      if (!cell || cell.formula !== null || cell.value.kind !== 'boolean') return;
      if (!wb.styles.get(sheet.effectiveStyleId(row, column)).checkbox) return;
      positions.push([row, column]);
      if (!cell.value.value) anyClear = true;
    };
    if (at) visit(at.row, at.column);
    else {
      // Walk whichever is smaller: the selected cells, or the cells the sheet stores.
      const ranges = this.ranges;
      const area = ranges.reduce((n, r) => n + r.cellCount, 0);
      if (area <= sheet.cellCount) {
        const seen = ranges.length > 1 ? new Set<number>() : null;
        for (const r of ranges)
          for (let row = r.top; row <= r.bottom; row++)
            for (let column = r.left; column <= r.right; column++) {
              if (seen) {
                const key = row * 16384 + column;
                if (seen.has(key)) continue;
                seen.add(key);
              }
              visit(row, column);
            }
      } else
        for (const [row, column] of sheet.enumerateCells())
          if (ranges.some((r) => r.contains(row, column))) visit(row, column);
    }
    if (positions.length === 0) return false;
    this.run(EditCommands.setChecked(this.sheetName, positions, anyClear));
    return true;
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

  /** Fill-handle drag: copy `source` across `target` (the source tiles), then select the target. */
  fillRange(source: GridRange, target: GridRange): void {
    const fill = new FillCommand(
      this.sheetName,
      CellRange.fromBounds(source.top, source.left, source.bottom, source.right),
      CellRange.fromBounds(target.top, target.left, target.bottom, target.right),
    );
    if (!this.#preflight(fill.projectWrites(this.session.workbook.peek()))) return;
    this.run(fill);
    this.selection.selectRange(target, { row: source.top, column: source.left });
  }

  /** The grid's hover tip for a cell: an error's reason first, else how to open its link. */
  hoverTip(row: number, column: number): string | null {
    const cell = this.sheet.find(row, column);
    if (!cell) return null;
    if (cell.value.kind === 'error') return `${cell.value.error.toString()} — ${cell.value.error.reason}`;
    if (cell.hyperlink)
      return `${navigator.platform.includes('Mac') ? '⌘' : 'Ctrl'}-click to open ${cell.hyperlink.target}`;
    return null;
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

  /**
   * ⌥↑/⌥↓ and ⌥←/⌥→: the rows (or columns) the selection is on move one step, swapping places with
   * the line next to them, and the selection goes with them.
   */
  moveLines(axis: 'row' | 'column', delta: -1 | 1): void {
    const r = this.range;
    const rows = axis === 'row';
    const command = new MoveLinesCommand(
      this.sheetName,
      axis,
      rows ? r.top : r.left,
      rows ? r.bottom : r.right,
      delta,
    );
    const refused = command.refusal(this.session.workbook.peek()) !== null;
    this.run(command);
    if (refused) return;
    const { row, column } = this.selection.active.peek();
    const dr = rows ? delta : 0;
    const dc = rows ? 0 : delta;
    this.selection.selectRange(
      { top: r.top + dr, left: r.left + dc, bottom: r.bottom + dr, right: r.right + dc },
      { row: row + dr, column: column + dc },
    );
  }

  async deleteRows(): Promise<void> {
    const r = this.range;
    const rows = new CellRange(
      new CellAddress(0, r.top),
      new CellAddress(SheetDimensions.maxColumns - 1, r.bottom),
    );
    const what = r.rowCount === 1 ? 'Delete this row' : `Delete ${r.rowCount} rows`;
    const asking = this.#largeRemoval([rows], what, 'Delete');
    if (asking && !(await asking)) return;
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

  async deleteColumns(): Promise<void> {
    const r = this.range;
    const columns = new CellRange(
      new CellAddress(r.left, 0),
      new CellAddress(r.right, SheetDimensions.maxRows - 1),
    );
    const what = r.columnCount === 1 ? 'Delete this column' : `Delete ${r.columnCount} columns`;
    const asking = this.#largeRemoval([columns], what, 'Delete');
    if (asking && !(await asking)) return;
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

  /**
   * Double-click on a column's edge: fit it, or, when it is one of the selected whole columns, every
   * selected whole column (⌘-clicked ones included).
   */
  autofitColumn(column: number): void {
    const whole = this.ranges.filter((r) => r.top === 0 && r.bottom >= MAX_ROW);
    if (!whole.some((r) => column >= r.left && column <= r.right)) {
      this.autofitColumns(CellRange.fromBounds(0, column, MAX_ROW, column));
      return;
    }
    this.autofitColumns(whole);
  }

  /** Fit each column the ranges touch (the selection's, by default) to its widest value in the column. */
  autofitColumns(ranges: CellRange | readonly CellRange[] = this.ranges): void {
    const spans = (Array.isArray(ranges) ? ranges : [ranges]) as readonly CellRange[];
    const touches = (column: number) => spans.some((r) => column >= r.left && column <= r.right);
    const ctx = document.createElement('canvas').getContext('2d')!;
    // The grid's own measure: its font, its padding, its box for turned text.
    const widths = fittedColumnWidths(
      ctx,
      this.session.workbook.peek(),
      this.sheet,
      (value, style) => this.session.display(value, style),
      { columns: touches },
    );
    const cmds = [...widths]
      .sort((a, b) => a[0] - b[0])
      .map(([c, width]) => EditCommands.setColumnWidth(this.sheetName, c, width));
    // Always "Autofitted columns …", even for one column, which on its own would read "Resized".
    if (cmds.length > 0) this.run(new CompositeEditCommand('Autofit Columns', cmds), spans);
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
    const hasHeader = this.range.isSingleCell && DataRegion.firstRowIsLabels(this.sheet, region);
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
    const hasHeader = DataRegion.firstRowIsLabels(this.sheet, region);
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
    return { range, hasHeader: DataRegion.firstRowIsLabels(this.sheet, range), table: null };
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
    this.run(SortFilterCommands.sortAndFilter(this.sheet, target, sort, filter, this.session.locale));
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
    if (sheet.activeFilter) return this.run(SortFilterCommands.removeAutoFilter(sheet, this.session.locale));
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
    this.run(SortFilterCommands.setAutoFilter(this.sheetName, range));
  }

  createTable(): void {
    const r = clampToUsed(this.sheet, this.regionOrSelection);
    if (!r) return this.info('Select some data to make a table.');
    // Every table gets a header row, so its columns always have filter and sort arrows.
    const wb = this.session.workbook.peek();
    const name = TableCommands.suggestName(wb, this.sheetName, r);
    const plan = TableCommands.create(wb, this.sheetName, r, { header: 'labels', name });
    if (!plan.ok) return this.info('That range overlaps an existing table.');
    this.run(plan.command);
    this.selection.select(plan.table.range.top, plan.table.range.left);
  }

  /** ⌘A: inside a table, the table first; pressed again (or outside one), the whole sheet. */
  /** Go To: the Name Box takes a cell, a range, `Sheet!A1` or a name (Enter goes, Esc returns). */
  focusNameBox(): void {
    this.nameBox.current?.focus();
  }

  selectAll(): void {
    const active = this.selection.active.peek();
    const t = this.activeTable;
    const s = this.selection.state.peek();
    const r = s.ranges.length === 1 ? s.ranges[0]! : null;
    const tr = t?.range;
    const onTable =
      !!tr && !!r && r.top === tr.top && r.left === tr.left && r.bottom === tr.bottom && r.right === tr.right;
    if (tr && !onTable)
      this.selection.selectRange({ top: tr.top, left: tr.left, bottom: tr.bottom, right: tr.right }, active);
    else this.selection.selectRange({ top: 0, left: 0, bottom: MAX_ROW, right: MAX_COLUMN }, active);
  }

  get activeTable(): Table | null {
    const { row, column } = this.selection.active.peek();
    return this.sheet.findTable(row, column) ?? null;
  }

  /** Rename the active table; shows why and returns false when the name can't be used. */
  renameTable(name: string): boolean {
    const t = this.activeTable;
    if (!t) return false;
    if (name.trim() === t.name) return true;
    const error = tableNameError(this.session.workbook.peek(), name.trim(), t);
    if (error) {
      this.error(error);
      return false;
    }
    this.run(DataEditCommands.renameTable(this.session.workbook.peek(), this.sheetName, t, name));
    return true;
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
    this.run(DataEditCommands.replaceValidation(this.sheetName, range, rule));
  }

  /** Items for a list rule's dropdown. */
  listItems(rule: DataValidation): string[] {
    return listItems(rule, this.resolver);
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

  /** Named ranges change the workbook directly (not undoable), then formulas are re-planned. */
  addNamedRange(nr: NamedRange): string | null {
    const wb = this.session.workbook.peek();
    const error = namedRangeError(wb, nr.name, nr.scope ?? null);
    if (error) return error;
    wb.addNamedRange(nr);
    this.session.namesChanged();
    return null;
  }

  removeNamedRange(nr: NamedRange): void {
    this.session.workbook.peek().removeNamedRange(nr.name, nr.scope ?? null);
    this.session.namesChanged();
  }

  /**
   * Clear every cell and reset the layout (line styles and hidden flags, filters, tables, validation,
   * merges, frozen panes) in one step. Asks first when that would discard more than 100 cells.
   */
  async resetSheetData(): Promise<void> {
    const sheet = this.sheet;
    const used = sheet.getUsedRange();
    if (!used) return this.info('The sheet is already empty.');
    if (sheet.cellCount > LARGE_REMOVAL_CELLS) {
      const message = `This clears ${sheet.cellCount.toLocaleString()} cells and the sheet’s layout. You can undo it.`;
      if (!(await this.confirm(`Reset “${this.sheetName}”?`, message, 'Reset'))) return;
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

  /**
   * "Pick on sheet" from the chart panel: show the data sheet with the field's current range
   * selected, and wait for Enter (`finishChartPick`) or Escape (`cancelChartPick`).
   */
  startChartPick(field: ChartField): void {
    const chartSheet = this.sheet;
    const chart = chartSheet.chart;
    if (!chartSheet.isChartSheet || !chart) return;
    const wb = this.session.workbook.peek();
    const target = ChartFields.pickTarget(wb, chart, field);
    const sheet = target ? wb.findSheet(target.sheet) : null;
    if (!target || !sheet) return this.info('There is no sheet with cells to pick.');
    this.session.setActiveSheet(wb.indexOf(sheet));
    const r = target.range;
    if (r) this.selection.selectRange({ top: r.top, left: r.left, bottom: r.bottom, right: r.right });
    this.chartPick.value = { chartSheet, field, label: ChartFields.label(field) };
    queueMicrotask(() => this.grid.current?.focus());
  }

  /** Enter while picking: put the selection in the chart's field and go back to the chart. */
  finishChartPick(): void {
    const pick = this.chartPick.peek();
    if (!pick) return;
    const reference = ChartFields.reference(this.sheetName, this.range);
    this.#endChartPick(pick);
    const chart = pick.chartSheet.chart;
    if (chart)
      this.run(
        ChartEditCommands.setChart(pick.chartSheet.name, ChartFields.set(chart, pick.field, reference)),
      );
  }

  /** Escape while picking: go back to the chart and change nothing. */
  cancelChartPick(): void {
    const pick = this.chartPick.peek();
    if (pick) this.#endChartPick(pick);
  }

  #endChartPick(pick: ChartPick): void {
    this.chartPick.value = null;
    const i = this.session.workbook.peek().indexOf(pick.chartSheet);
    if (i >= 0) this.session.setActiveSheet(i);
  }

  // ── Clipboard ────────────────────────────────────────────────────────────

  /** Copy to our block plus the system clipboard as TSV of displayed values. */
  copy(cut = false, data?: DataTransfer | null): void {
    const block = ClipboardBlock.capture(this.sheet, this.range, cut);
    const text = this.blockText(block);
    this.#clipboard = { block, text, generation: this.session.generation };
    this.marquee.value = { sheet: this.sheetName, range: this.selection.range.peek(), cut };
    if (data) data.setData('text/plain', text);
    else void navigator.clipboard?.writeText(text).catch(() => {});
    if (cut) this.info('Cut — paste to move.');
  }

  /** Paste: our own block when the system clipboard still holds what we copied, else TSV text. */
  paste(text: string | null, mode: 'all' | 'values' = 'all'): void {
    let own = this.#clipboard;
    // A cut belongs to the document it came from: moving it into another one would clear cells there.
    if (own?.block.isCut && own.generation !== this.session.generation) {
      this.#clipboard = own = null;
      this.marquee.value = null;
    }
    const sheet = this.sheetName;
    if (own && (text === null || normalize(text) === normalize(own.text))) {
      const cut = own.block.isCut;
      const paste = new PasteBlockCommand(sheet, this.range, own.block, mode);
      if (!this.#preflight(paste.projectWrites())) return;
      this.run(paste);
      if (cut) {
        this.#clipboard = null;
        this.marquee.value = null;
      }
      return;
    }
    if (!text) return;
    // Spreadsheets put TSV on the clipboard. Text without a tab is lines of plain text (a comma is
    // just a comma), unless it is one quoted cell, which is how a cell with a line break copies.
    const body = text.replace(/\r\n?/g, '\n').replace(/\n$/, '');
    const rows =
      body.includes('\t') || /^"(?:[^"]|"")*"$/.test(body)
        ? parseCsv(body, { delimiter: '\t' })
        : body.split('\n').map((line) => [line]);
    const locale = this.session.locale;
    const { row, column } = this.selection.active.peek();
    const textWidth = rows.reduce((w, r) => Math.max(w, r.length), 0);
    if (!PasteBlockCommand.fits(row, column, rows.length, textWidth)) {
      this.#showToast(PasteBlockCommand.doesNotFit, 'error');
      return;
    }
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
    const width = textWidth;
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

  /** Every cell on the active sheet whose shown text or formula contains the query, row-major. */
  findAll(query: string): FindMatch[] {
    const wb = this.session.workbook.peek();
    const sheet = this.sheet;
    return findMatches(sheet, query, (r, c, cell) =>
      this.session.display(cell.value, wb.styles.get(sheet.effectiveStyleId(r, c))),
    );
  }

  /** What the cell at `row`, `column` on the active sheet shows, as the grid draws it. */
  shownText(row: number, column: number): string {
    const sheet = this.sheet;
    const cell = sheet.find(row, column);
    if (!cell) return '';
    const style = this.session.workbook.peek().styles.get(sheet.effectiveStyleId(row, column));
    return this.session.display(cell.value, style);
  }

  /** Select a cell Find matched and scroll it into view. */
  goToMatch(match: { readonly row: number; readonly column: number }): void {
    this.selection.select(match.row, match.column);
    this.grid.current?.revealActive();
  }

  /** Next cell (row-major, wrapping) whose shown text or formula contains the query. */
  findNext(query: string, backwards = false): boolean {
    const hit = nextMatch(this.findAll(query), this.selection.active.peek(), backwards);
    if (!hit) return false;
    this.goToMatch(hit);
    return true;
  }

  /** The edit that writes `next` (a cell's replaced text or formula) back into it. */
  #replaceWrite(row: number, column: number, cell: Cell, next: string) {
    const s = this.session;
    if (cell.formula === null) {
      // Text stays text, as in Calc: '007 doesn't turn into 7, nor 1/2 into a date.
      const value = next === '' ? CellValue.empty : CellValue.text(next);
      return {
        write: { row, column, value, isFormula: false },
        command: EditCommands.setValue(this.sheetName, row, column, value),
      };
    }
    const parsed = InputParser.parse(next, s.locale);
    return {
      write: { row, column, value: parsed.value, isFormula: parsed.formula !== null },
      command: EditCommands.setInput(this.sheetName, row, column, next, s.locale, s.workbook.peek()),
    };
  }

  /**
   * Replace: when the active cell holds the query, replace it there (one undo step), then go to
   * the next match. When it doesn't, only go to the next match, so the first press shows what the
   * second will change. `replaced` says whether a cell changed, `found` whether any match is left.
   */
  replaceNext(query: string, replacement: string): { replaced: boolean; found: boolean } {
    if (!query) return { replaced: false, found: false };
    const { row, column } = this.selection.active.peek();
    const cell = this.sheet.find(row, column);
    const text = replaceableText(cell);
    const next = text === null ? null : replaceIn(text, query, replacement);
    let replaced = false;
    if (cell && next !== null) {
      const { write, command } = this.#replaceWrite(row, column, cell, next);
      if (!this.#preflight([write])) return { replaced: false, found: true };
      this.run(each('Replace', [command]));
      replaced = true;
    }
    return { replaced, found: this.findNext(query) };
  }

  replaceAll(query: string, replacement: string): number {
    if (!query) return 0;
    const cmds: EditCommand[] = [];
    const writes: { row: number; column: number; value: CellValue; isFormula: boolean }[] = [];
    // Snapshot the stored cells first: the commands run later, but the walk must not see its own writes.
    for (const [r, c, cell] of [...this.sheet.enumerateCells()]) {
      const text = replaceableText(cell);
      const next = text === null ? null : replaceIn(text, query, replacement);
      if (next === null) continue;
      const { write, command } = this.#replaceWrite(r, c, cell, next);
      writes.push(write);
      cmds.push(command);
    }
    if (!this.#preflight(writes)) return 0;
    this.run(each('Replace All', cmds));
    return cmds.length;
  }

  /**
   * A named range (the sheet's own first, then the workbook's) or a table as a `Sheet!A1:B2`
   * reference, or null when `text` names neither. Cell references win: `A1` is never looked up as a name.
   */
  #namedTarget(text: string): string | null {
    if (text.includes('!') || CellRange.tryParse(text) || CellAddress.tryParse(text)) return null;
    const wb = this.session.workbook.peek();
    const range = wb.findNamedRange(text, this.sheetName)?.target;
    if (!range) {
      const found = findTableByName(wb, text);
      if (!found) return null;
      return `'${found.sheet.name.replaceAll("'", "''")}'!${found.table.range.withSheet(null).toString()}`;
    }
    const sheet = range.start.sheet ?? this.sheetName;
    const bounds = CellRange.fromBounds(range.top, range.left, range.bottom, range.right);
    return `'${sheet.replaceAll("'", "''")}'!${bounds.toString()}`;
  }

  /** Name box: `B7`, `B2:D9`, `Sheet2!A1`, a named range or a table. False when it doesn't parse. */
  goTo(reference: string): boolean {
    const text = reference.trim();
    const named = this.#namedTarget(text);
    if (named) return this.goTo(named);
    const bang = text.lastIndexOf('!');
    if (bang > 0) {
      const name = text.slice(0, bang).replace(/^'|'$/g, '').replaceAll("''", "'");
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
    this.grid.current?.revealActive();
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

  async deleteSheet(index = this.session.activeIndex.peek()): Promise<void> {
    const names = this.session.sheetNames.peek();
    if (names.length <= 1) return this.info('A workbook needs at least one sheet.');
    const name = names[index];
    if (name === undefined) return;
    // PRD §8.7: ask only when something would be lost or broken; an empty, unused sheet just goes
    // (undo still brings it back).
    const message = deleteSheetMessage(name, sheetDeletionImpact(this.session.workbook.peek(), name));
    if (message !== null && !(await this.confirm('Delete sheet?', message, 'Delete'))) return;
    this.session.deleteSheet(index);
  }
}

// ── helpers ──────────────────────────────────────────────────────────────

/** The question to ask before deleting sheet `name`, or null when nothing would be lost or broken. */
export function deleteSheetMessage(name: string, impact: SheetDeletionImpact): string | null {
  const { empty, formulas, otherReferences } = impact;
  if (empty && formulas === 0 && otherReferences === 0) return null;
  const parts = [`Delete the sheet “${name}”?`];
  if (formulas > 0)
    parts.push(
      formulas === 1
        ? '1 formula on another sheet refers to it and will show #REF!.'
        : `${formulas} formulas on other sheets refer to it and will show #REF!.`,
    );
  if (otherReferences > 0)
    parts.push(
      otherReferences === 1
        ? '1 chart or validation list uses its cells and will lose them.'
        : `${otherReferences} charts or validation lists use its cells and will lose them.`,
    );
  return parts.join(' ');
}

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

function normalize(text: string): string {
  return text.replace(/\r\n?/g, '\n').replace(/\n$/, '');
}
