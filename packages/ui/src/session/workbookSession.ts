import { batch, computed, signal, type ReadonlySignal } from '@preact/signals';
import {
  CsvOptions,
  EditHistory,
  NumberFormatter,
  NumberFormats,
  OdsReader,
  OdsWriter,
  RecalculationEngine,
  Workbook,
  readCsv,
  writeCsv,
  type CellStyle,
  type CellValue,
  type EditCommand,
  type EditResult,
  type OdsPackage,
  type Worksheet,
} from '@cellestra/core';
import type { DocumentRef } from '../platform/fileStore.ts';
import type { SelectionState } from './selection.ts';

export type DocumentFormat = 'ods' | 'csv';

/** What a step of history restores the selection to. */
interface HistoryEntry {
  readonly sheet: string;
  readonly selection: SelectionState | null;
}

/**
 * One open document: the workbook, its recalc engine and undo history, the active sheet, and the
 * dirty flag. The UI reads it through signals; every model change bumps `version` so the grid and
 * chrome repaint. Edits go through `execute` so they are undoable and recalculated.
 */
export class WorkbookSession {
  readonly #workbook = signal<Workbook>(Workbook.createNew());
  readonly #activeIndex = signal(0);
  readonly #version = signal(0);
  readonly #structure = signal(0);
  readonly #modified = signal(false);
  readonly #document = signal<DocumentRef | null>(null);
  readonly #format = signal<DocumentFormat>('ods');
  readonly #circular = signal(false);

  #engine: RecalculationEngine;
  #history: EditHistory;
  #package: OdsPackage | null = null;
  readonly #undoSelections: HistoryEntry[] = [];
  readonly #redoSelections: HistoryEntry[] = [];

  /** Set by the app so undo can return the cursor to where an edit happened. */
  captureSelection: () => SelectionState | null = () => null;
  restoreSelection: (s: SelectionState) => void = () => {};

  /** Bumps on every change the grid must repaint for. */
  readonly version: ReadonlySignal<number> = this.#version;
  /** Bumps when sheets are added, removed, renamed, reordered, or the workbook is replaced. */
  readonly structure: ReadonlySignal<number> = this.#structure;
  readonly modified: ReadonlySignal<boolean> = this.#modified;
  readonly document: ReadonlySignal<DocumentRef | null> = this.#document;
  readonly format: ReadonlySignal<DocumentFormat> = this.#format;
  readonly hasCircularReference: ReadonlySignal<boolean> = this.#circular;

  readonly workbook: ReadonlySignal<Workbook> = this.#workbook;
  readonly activeSheet: ReadonlySignal<Worksheet> = computed(() => {
    void this.#structure.value;
    const sheets = this.#workbook.value.worksheets;
    return sheets[Math.min(this.#activeIndex.value, sheets.length - 1)]!;
  });
  readonly sheetNames: ReadonlySignal<readonly string[]> = computed(() => {
    void this.#structure.value;
    return this.#workbook.value.worksheets.map((s) => s.name);
  });
  readonly activeIndex: ReadonlySignal<number> = this.#activeIndex;
  readonly title: ReadonlySignal<string> = computed(() => this.#document.value?.name ?? 'Untitled.ods');
  readonly canUndo: ReadonlySignal<boolean> = computed(() => (this.#version.value, this.#history.canUndo));
  readonly canRedo: ReadonlySignal<boolean> = computed(() => (this.#version.value, this.#history.canRedo));

  constructor() {
    const wb = this.#workbook.peek();
    this.#engine = new RecalculationEngine(wb);
    this.#history = new EditHistory(wb);
  }

  get locale(): string {
    return this.#workbook.peek().settings.effectiveLocale;
  }

  get undoLabel(): string | null {
    return this.#history.undoLabel;
  }

  get redoLabel(): string | null {
    return this.#history.redoLabel;
  }

  // ── Documents ────────────────────────────────────────────────────────────

  newDocument(): void {
    this.#adopt(Workbook.createNew(), null, null, 'ods');
  }

  /** Parse bytes by file name (.ods, else CSV/TSV) and make them the open document. Throws on bad files. */
  load(bytes: Uint8Array, ref: DocumentRef | null, name = ref?.name ?? 'Untitled.ods'): void {
    const format = formatOf(name);
    if (format === 'ods') {
      const { workbook, package: pkg } = OdsReader.readWithPackage(bytes);
      this.#adopt(workbook, pkg, ref, 'ods');
    } else {
      const text = new TextDecoder('utf-8').decode(bytes).replace(/^\uFEFF/, '');
      const options = /\.tsv$/i.test(name) ? CsvOptions.tab : sniffDelimiter(text);
      this.#adopt(readCsv(text, options), null, ref, 'csv');
    }
  }

  /** Bytes in the document's own format (CSV writes the active sheet's displayed values). */
  serialize(format: DocumentFormat = this.#format.peek(), name = this.title.peek()): Uint8Array {
    const wb = this.#workbook.peek();
    if (format === 'ods') return OdsWriter.write(wb, this.#package);
    const options = /\.tsv$/i.test(name) ? CsvOptions.tab : CsvOptions.comma;
    return new TextEncoder().encode(writeCsv(wb, this.activeSheet.peek(), options));
  }

  /** Record a successful save: the new ref (name, version) and a clean flag. */
  markSaved(ref: DocumentRef): void {
    batch(() => {
      this.#document.value = ref;
      this.#format.value = formatOf(ref.name);
      this.#modified.value = false;
    });
  }

  /** Replace the document with a reloaded copy (external change). Keeps the active sheet when it still exists. */
  reload(bytes: Uint8Array, ref: DocumentRef): void {
    const name = this.activeSheet.peek().name;
    this.load(bytes, ref);
    const i = this.#workbook.peek().worksheets.findIndex((s) => s.name === name);
    if (i >= 0) this.setActiveSheet(i);
  }

  #adopt(wb: Workbook, pkg: OdsPackage | null, ref: DocumentRef | null, format: DocumentFormat): void {
    this.#engine = new RecalculationEngine(wb);
    this.#engine.rebuildGraph();
    const result = this.#engine.recalculateAll();
    this.#history = new EditHistory(wb);
    this.#package = pkg;
    this.#undoSelections.length = 0;
    this.#redoSelections.length = 0;
    batch(() => {
      this.#workbook.value = wb;
      this.#activeIndex.value = Math.min(Math.max(wb.settings.activeSheetIndex, 0), wb.sheetCount - 1);
      this.#document.value = ref;
      this.#format.value = format;
      this.#modified.value = false;
      this.#circular.value = result.hasCircularReference;
      this.#structure.value++;
      this.#version.value++;
    });
  }

  // ── Edits ────────────────────────────────────────────────────────────────

  execute(command: EditCommand): void {
    this.#undoSelections.push(this.#entry());
    if (this.#undoSelections.length > EditHistory.maxDepth) this.#undoSelections.shift();
    this.#redoSelections.length = 0;
    this.#apply(this.#history.execute(command));
  }

  undo(): void {
    if (!this.#history.canUndo) return;
    this.#redoSelections.push(this.#entry());
    this.#apply(this.#history.undo());
    this.#restore(this.#undoSelections.pop());
  }

  redo(): void {
    if (!this.#history.canRedo) return;
    this.#undoSelections.push(this.#entry());
    this.#apply(this.#history.redo());
    this.#restore(this.#redoSelections.pop());
  }

  recalculateAll(): void {
    this.#engine.rebuildGraph();
    this.#circular.value = this.#engine.recalculateAll().hasCircularReference;
    this.#version.value++;
  }

  /** Recalculate TODAY/NOW/RAND cells (on focus or a timer). */
  recalculateVolatile(): void {
    if (this.#engine.volatileCells.length === 0) return;
    this.#engine.recalculateVolatile();
    this.#version.value++;
  }

  /** A change made outside the history (resizing columns isn't worth an undo step in some flows). */
  touch(structural = false): void {
    batch(() => {
      this.#modified.value = true;
      if (structural) this.#structure.value++;
      this.#version.value++;
    });
  }

  #apply(result: EditResult): void {
    let circular = this.#circular.peek();
    if (result.structural) {
      this.#engine.rebuildGraph();
      circular = this.#engine.recalculateAll().hasCircularReference;
    } else if (result.changedCells.length > 0) {
      circular = this.#engine.recordEdits(result.changedCells).hasCircularReference;
    }
    batch(() => {
      this.#modified.value = true;
      this.#circular.value = circular;
      if (result.structural) this.#structure.value++;
      this.#version.value++;
    });
  }

  #entry(): HistoryEntry {
    return { sheet: this.activeSheet.peek().name, selection: this.captureSelection() };
  }

  #restore(entry: HistoryEntry | undefined): void {
    if (!entry) return;
    const i = this.#workbook.peek().worksheets.findIndex((s) => s.name === entry.sheet);
    if (i >= 0) this.setActiveSheet(i);
    if (entry.selection) this.restoreSelection(entry.selection);
  }

  // ── Sheets (not undoable, PRD §8.5: they clear the history) ─────────────

  setActiveSheet(index: number): void {
    const wb = this.#workbook.peek();
    if (index < 0 || index >= wb.sheetCount) return;
    wb.settings.activeSheetIndex = index;
    this.#activeIndex.value = index;
  }

  addSheet(): void {
    const wb = this.#workbook.peek();
    wb.addSheet();
    this.#structural();
    this.setActiveSheet(wb.sheetCount - 1);
  }

  deleteSheet(index: number): void {
    const wb = this.#workbook.peek();
    const sheet = wb.worksheets[index];
    if (!sheet || wb.sheetCount <= 1 || !wb.removeSheet(sheet)) return;
    this.#structural();
    this.setActiveSheet(Math.min(index, wb.sheetCount - 1));
  }

  duplicateSheet(index: number): void {
    const wb = this.#workbook.peek();
    const sheet = wb.worksheets[index];
    if (!sheet) return;
    const copy = wb.duplicateSheet(sheet);
    this.#structural();
    this.setActiveSheet(wb.indexOf(copy));
  }

  /** Add a sheet built by `create` (e.g. a chart sheet); like every sheet-list change, clears history. */
  insertSheet(create: (workbook: Workbook) => Worksheet): void {
    const wb = this.#workbook.peek();
    const sheet = create(wb);
    this.#structural();
    this.setActiveSheet(wb.indexOf(sheet));
  }

  /** Rename; returns an error message for invalid or duplicate names. */
  renameSheet(index: number, name: string): string | null {
    const wb = this.#workbook.peek();
    const sheet = wb.worksheets[index];
    if (!sheet) return null;
    if (sheet.name === name) return null;
    try {
      wb.renameSheet(sheet, name);
    } catch (e) {
      return e instanceof Error ? e.message : String(e);
    }
    this.#structural();
    return null;
  }

  moveSheet(from: number, to: number): void {
    const wb = this.#workbook.peek();
    const active = this.activeSheet.peek();
    wb.moveSheet(from, to);
    this.#structural();
    this.setActiveSheet(wb.indexOf(active));
  }

  #structural(): void {
    this.#history.clear();
    this.#undoSelections.length = 0;
    this.#redoSelections.length = 0;
    this.#engine.rebuildGraph();
    const result = this.#engine.recalculateAll();
    batch(() => {
      this.#circular.value = result.hasCircularReference;
      this.#modified.value = true;
      this.#structure.value++;
      this.#version.value++;
    });
  }

  // ── Display ──────────────────────────────────────────────────────────────

  /** A value as the grid and status bar show it, under its style's number format. */
  display(value: CellValue, style: CellStyle): string {
    return NumberFormatter.format(value, style.numberFormat.code, this.locale);
  }

  /** What the formula bar and inline editor start with for a cell. */
  editText(row: number, column: number): string {
    const sheet = this.activeSheet.peek();
    const cell = sheet.find(row, column);
    if (!cell) return '';
    if (cell.formula !== null) return '=' + cell.formula;
    const v = cell.value;
    switch (v.kind) {
      case 'empty':
        return '';
      case 'boolean':
        return v.value ? 'TRUE' : 'FALSE';
      case 'error':
        return v.error.toString();
      case 'text':
        // Text that would re-parse as something else keeps its apostrophe flag.
        return /^[=+\-']|^(true|false)$/i.test(v.value) || looksNumeric(v.value) ? "'" + v.value : v.value;
      case 'number': {
        const style = this.#workbook.peek().styles.get(sheet.effectiveStyleId(row, column));
        const kind = NumberFormats.kind(style.numberFormat);
        if (kind === 'date' || kind === 'time' || kind === 'dateTime')
          return NumberFormatter.format(v, editCodeFor(kind), this.locale);
        if (kind === 'percent') return NumberFormatter.formatGeneral(v.value * 100, this.locale) + '%';
        return NumberFormatter.formatGeneral(v.value, this.locale);
      }
    }
  }
}

function editCodeFor(kind: string): string {
  if (kind === 'time') return 'HH:MM:SS';
  if (kind === 'dateTime') return 'YYYY-MM-DD HH:MM:SS';
  return 'YYYY-MM-DD';
}

function looksNumeric(s: string): boolean {
  return /^\s*[-+]?(\d[\d,.\s]*|\.\d+)(e[-+]?\d+)?%?\s*$/i.test(s);
}

export function formatOf(name: string): DocumentFormat {
  return /\.(csv|tsv|txt)$/i.test(name) ? 'csv' : 'ods';
}

/** Comma unless the first line has more semicolons or tabs. */
function sniffDelimiter(text: string): CsvOptions {
  const nl = text.indexOf('\n');
  const line = nl < 0 ? text : text.slice(0, nl);
  const count = (ch: string): number => line.split(ch).length - 1;
  const comma = count(',');
  const semi = count(';');
  const tab = count('\t');
  if (tab > comma && tab >= semi) return CsvOptions.tab;
  if (semi > comma) return CsvOptions.semicolon;
  return CsvOptions.comma;
}
