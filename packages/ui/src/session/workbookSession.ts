import { batch, computed, signal, type ReadonlySignal } from '@preact/signals';
import {
  CsvOptions,
  EditHistory,
  EditRefusedError,
  NumberFormatter,
  NumberFormats,
  OdsReader,
  OdsWriter,
  RecalculationEngine,
  DeleteSheetCommand,
  RenameSheetCommand,
  SheetName,
  Workbook,
  readCsv,
  writeCsv,
  type CellStyle,
  type CalculationMode,
  type CellValue,
  type EditCommand,
  EditResult,
  type OdsPackage,
  type Worksheet,
} from '@cellestra/core';
import type { DocumentRef } from '../platform/fileStore.ts';
import type { GridRange, SelectionState } from './selection.ts';
import { narrateEdit, narrateHistory } from './statusText.ts';

/** A cell's position within its sheet. */
export interface CellPosition {
  readonly row: number;
  readonly column: number;
}
/** How many versions back `cellsChangedSince` can answer, and the most cells one version logs. */
const CELL_LOG_VERSIONS = 64;
const CELL_LOG_MAX_CELLS = 50_000;

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
  readonly #lastAction = signal('');
  readonly #savedAt = signal<Date | null>(null);

  #engine: RecalculationEngine;
  #history: EditHistory;
  #package: OdsPackage | null = null;
  /** Counts every dirtying change, so a save can tell whether edits landed while it was in flight. */
  #edits = 0;
  #generation = 0;
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
  /** The status bar's narration of the most recent edit, undo or redo ("Pasted A1:C4"). */
  readonly lastAction: ReadonlySignal<string> = this.#lastAction;
  /** When this document was last saved in this session, or null. */
  readonly savedAt: ReadonlySignal<Date | null> = this.#savedAt;
  /** The clock `savedAt` reads (tests set it). */
  now: () => Date = () => new Date();

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
  readonly calculationMode: ReadonlySignal<CalculationMode> = computed(
    () => (this.#version.value, this.#workbook.value.settings.calculationMode),
  );

  constructor() {
    const wb = this.#workbook.peek();
    this.#engine = new RecalculationEngine(wb);
    this.#history = new EditHistory(wb);
  }

  /** Bumps on every change that dirties the document; pass it to `markSaved` from serialize time. */
  get editCount(): number {
    return this.#edits;
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

  /**
   * Parse bytes by file name (.ods, else CSV/TSV) and make them the open document. Throws on bad
   * files. `saveAs` overrides the format Save writes (recovery restores ODS bytes of a CSV document).
   */
  load(
    bytes: Uint8Array,
    ref: DocumentRef | null,
    name = ref?.name ?? 'Untitled.ods',
    saveAs?: DocumentFormat,
  ): void {
    const format = formatOf(name);
    if (format === 'ods') {
      const { workbook, package: pkg } = OdsReader.readWithPackage(bytes);
      this.#adopt(workbook, pkg, ref, saveAs ?? 'ods');
    } else {
      const text = new TextDecoder('utf-8').decode(bytes).replace(/^\uFEFF/, '');
      const options = /\.tsv$/i.test(name) ? CsvOptions.tab : sniffDelimiter(text);
      this.#adopt(readCsv(text, options), null, ref, 'csv');
    }
  }

  /**
   * The sheet CSV writes. A CSV document is its first grid sheet, whichever sheet is on screen, so
   * Save never swaps the file's data for a scratch sheet or a chart. An export of an .ods writes
   * the active sheet, unless that is a chart.
   */
  csvSheet(format: DocumentFormat = this.#format.peek()): Worksheet {
    const wb = this.#workbook.peek();
    const active = this.activeSheet.peek();
    if (format !== 'csv' && !active.isChartSheet) return active;
    return wb.worksheets.find((s) => !s.isChartSheet) ?? active;
  }

  /** Bytes in the document's own format (CSV writes one sheet's displayed values: `csvSheet`). */
  serialize(format: DocumentFormat = this.#format.peek(), name = this.title.peek()): Uint8Array {
    const wb = this.#workbook.peek();
    if (format === 'ods') return OdsWriter.write(wb, this.#package);
    const options = /\.tsv$/i.test(name) ? CsvOptions.tab : CsvOptions.comma;
    return new TextEncoder().encode(writeCsv(wb, this.csvSheet(), options));
  }

  /**
   * Record a successful save: the new ref (name, version), and a clean flag unless edits landed
   * after `editCount` (the value of `editCount` when the saved bytes were serialized).
   */
  markSaved(ref: DocumentRef, editCount = this.#edits): void {
    batch(() => {
      this.#document.value = ref;
      this.#format.value = formatOf(ref.name);
      if (editCount === this.#edits) this.#modified.value = false;
      this.#savedAt.value = this.now();
    });
  }

  /** Replace the document with a reloaded copy (external change). Keeps the active sheet when it still exists. */
  reload(bytes: Uint8Array, ref: DocumentRef): void {
    const name = this.activeSheet.peek().name;
    this.load(bytes, ref);
    const i = this.#workbook.peek().worksheets.findIndex((s) => s.name === name);
    if (i >= 0) this.setActiveSheet(i);
  }

  /**
   * Bumped whenever another document (or a reloaded copy) replaces the open one. A save started
   * before that must not stamp its file onto the new document.
   */
  get generation(): number {
    return this.#generation;
  }

  #adopt(wb: Workbook, pkg: OdsPackage | null, ref: DocumentRef | null, format: DocumentFormat): void {
    this.#generation++;
    this.#engine = new RecalculationEngine(wb);
    this.#engine.rebuildGraph();
    // PRD §8.6: manual calculation keeps the file's cached values until the user recalculates.
    const manual = wb.settings.calculationMode === 'manual';
    const circular = manual ? false : this.#engine.recalculateAll().hasCircularReference;
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
      this.#circular.value = circular;
      this.#lastAction.value = '';
      this.#savedAt.value = null;
      this.#structure.value++;
      this.#version.value++;
    });
  }

  // ── Edits ────────────────────────────────────────────────────────────────

  /** Told when an edit is refused (a sort over merged cells, an insert that would push data off the sheet). */
  onRefused: (message: string) => void = () => {};

  /** `target` names where the edit landed in the status line when it reports no cells and the
   * selection isn't it (a column fitted by double-clicking its edge). */
  execute(command: EditCommand, target?: GridRange | readonly GridRange[]): void {
    const entry = this.#entry();
    // A command that throws (a refused edit) leaves the history as it was, so the selection
    // stacks only move once it has run.
    let result: EditResult;
    try {
      result = this.#history.execute(command);
    } catch (e) {
      if (!(e instanceof EditRefusedError)) throw e;
      this.onRefused(e.message);
      return;
    }
    this.#undoSelections.push(entry);
    if (this.#undoSelections.length > EditHistory.maxDepth) this.#undoSelections.shift();
    this.#redoSelections.length = 0;
    this.#apply(result);
    const ranges = entry.selection?.ranges;
    this.#lastAction.value = narrateEdit(
      command.label,
      result,
      target ?? ranges?.[ranges.length - 1] ?? null,
    );
  }

  undo(): void {
    const label = this.#history.undoLabel;
    if (label === null) return;
    this.#redoSelections.push(this.#entry());
    this.#apply(this.#history.undo());
    this.#restore(this.#undoSelections.pop());
    this.#lastAction.value = narrateHistory('undo', label);
  }

  redo(): void {
    const label = this.#history.redoLabel;
    if (label === null) return;
    this.#undoSelections.push(this.#entry());
    this.#apply(this.#history.redo());
    this.#restore(this.#redoSelections.pop());
    this.#lastAction.value = narrateHistory('redo', label);
  }

  /** Recalculate every formula (F9), in either calculation mode. */
  recalculateAll(): void {
    this.#engine.rebuildGraph();
    const circular = this.#engine.recalculateAll().hasCircularReference;
    batch(() => {
      this.#circular.value = circular;
      this.#lastAction.value = 'Recalculated';
      this.#version.value++;
    });
  }

  /**
   * Switch between automatic and manual calculation (saved with the file, so it dirties it). Going
   * back to automatic recalculates at once, so what the sheet shows is current again.
   */
  setCalculationMode(mode: CalculationMode): void {
    const wb = this.#workbook.peek();
    if (wb.settings.calculationMode === mode) return;
    wb.settings.calculationMode = mode;
    let circular = this.#circular.peek();
    if (mode === 'automatic') {
      this.#engine.rebuildGraph();
      circular = this.#engine.recalculateAll().hasCircularReference;
    }
    batch(() => {
      this.#circular.value = circular;
      this.#modified.value = true;
      this.#edits++;
      this.#lastAction.value = `Calculation: ${mode === 'manual' ? 'Manual' : 'Automatic'}`;
      this.#version.value++;
    });
  }

  /** Recalculate TODAY/NOW/RAND cells (on focus or a timer). Manual calculation waits for F9. */
  recalculateVolatile(): void {
    if (this.#workbook.peek().settings.calculationMode === 'manual') return;
    if (this.#engine.volatileCells.length === 0) return;
    const recalc = this.#engine.recalculateVolatile();
    this.#logCells([recalc.evaluatedCells]);
    this.#version.value++;
  }

  /**
   * Named ranges changed outside the history: re-plan every formula from the new names and mark the
   * document modified. Manual calculation leaves the results for F9.
   */
  namesChanged(): void {
    const recalc = this.#engine.apply(EditResult.structuralChange);
    if (recalc.calculated) this.#circular.value = recalc.hasCircularReference;
    this.touch();
  }

  /** A change made outside the history (resizing columns isn't worth an undo step in some flows). */
  touch(structural = false): void {
    batch(() => {
      this.#modified.value = true;
      this.#edits++;
      if (structural) this.#structure.value++;
      this.#version.value++;
    });
  }

  #apply(result: EditResult): void {
    const recalc = this.#engine.apply(result);
    // Manual calculation evaluates nothing, so it can't say whether a cycle is still there.
    const circular = recalc.calculated ? recalc.hasCircularReference : this.#circular.peek();
    const evaluated = result.structural ? [] : recalc.evaluatedCells;
    // An edit that reports neither cells nor lines changed something unknown here: logged as such.
    // Logged before the version moves, because a view redraws as soon as it does.
    const lines = result.changedRows.length + result.changedColumns.length > 0;
    this.#logCells(
      result.structural || (result.changedCells.length === 0 && !lines)
        ? null
        : lines
          ? this.#cellsInLines(result, [result.changedCells, evaluated])
          : [result.changedCells, evaluated],
    );
    batch(() => {
      this.#modified.value = true;
      this.#edits++;
      this.#circular.value = circular;
      if (result.structural) this.#structure.value++;
      this.#version.value++;
    });
  }

  /**
   * `groups` plus the stored cells in the rows and columns the edit reported (a line style, a width,
   * a merge), so a view re-measures those cells only. Null when they are too many to log.
   */
  #cellsInLines(
    result: EditResult,
    groups: (readonly (CellPosition & { readonly sheet: string })[])[],
  ): (readonly (CellPosition & { readonly sheet: string })[])[] | null {
    const wb = this.#workbook.peek();
    let total = 0;
    for (const g of groups) total += g.length;
    const names = new Set(
      [...result.changedRows, ...result.changedColumns].map((l) => l.sheet.toLowerCase()),
    );
    for (const name of names) {
      const sheet = wb.worksheets.find((w) => w.name.toLowerCase() === name);
      if (!sheet) continue;
      const rows = result.changedRows.filter((l) => l.sheet.toLowerCase() === name);
      const columns = result.changedColumns.filter((l) => l.sheet.toLowerCase() === name);
      const within = (lines: typeof rows, i: number) => lines.some((l) => i >= l.first && i <= l.last);
      const cells: (CellPosition & { readonly sheet: string })[] = [];
      for (const [row, column] of sheet.enumerateCells()) {
        if (!within(rows, row) && !within(columns, column)) continue;
        if (++total > CELL_LOG_MAX_CELLS) return null;
        cells.push({ sheet: sheet.name, row, column });
      }
      groups.push(cells);
    }
    return groups;
  }

  /** Cells each recent version changed, per lower-cased sheet name; null when that isn't known. */
  readonly #cellLog = new Map<number, Map<string, CellPosition[]> | null>();

  /** Log what the next version will change. Call it just before bumping the version. */
  #logCells(groups: (readonly (CellPosition & { readonly sheet: string })[])[] | null): void {
    const version = this.#version.peek() + 1;
    if (this.#cellLog.size >= CELL_LOG_VERSIONS)
      for (const v of this.#cellLog.keys()) if (v <= version - CELL_LOG_VERSIONS) this.#cellLog.delete(v);
    let total = 0;
    for (const g of groups ?? []) total += g.length;
    if (!groups || total > CELL_LOG_MAX_CELLS) {
      this.#cellLog.set(version, null);
      return;
    }
    const bySheet = new Map<string, CellPosition[]>();
    for (const g of groups)
      for (const c of g) {
        const key = c.sheet.toLowerCase();
        let list = bySheet.get(key);
        if (!list) bySheet.set(key, (list = []));
        list.push({ row: c.row, column: c.column });
      }
    this.#cellLog.set(version, bySheet);
  }

  /**
   * The cells of `sheet` changed (edited or recalculated) since `version`, so a view can update
   * only what moved. Null when that can't be told (a structural or style edit, a load, or too long
   * ago): the view then rebuilds from scratch.
   */
  cellsChangedSince(sheet: string, version: number): CellPosition[] | null {
    const now = this.#version.peek();
    if (version > now || now - version >= CELL_LOG_VERSIONS) return null;
    const key = sheet.toLowerCase();
    const out: CellPosition[] = [];
    for (let v = version + 1; v <= now; v++) {
      const entry = this.#cellLog.get(v);
      if (!entry) return null;
      const cells = entry.get(key);
      if (cells) for (const c of cells) out.push(c);
    }
    return out;
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

  // ── Sheets (PRD §8.5: these clear the history; only rename and delete are undoable) ──

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

  /**
   * Delete a sheet as an undoable edit: references to it elsewhere become #REF!, so a sheet added
   * later under the same name can't silently pick them up.
   */
  deleteSheet(index: number): void {
    const wb = this.#workbook.peek();
    const sheet = wb.worksheets[index];
    if (!sheet || wb.sheetCount <= 1) return;
    this.execute(new DeleteSheetCommand(sheet.name));
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

  /**
   * Rename, rewriting every reference to the sheet; unlike the other sheet changes this is undoable.
   * Returns an error message for invalid or duplicate names.
   */
  renameSheet(index: number, name: string): string | null {
    const wb = this.#workbook.peek();
    const sheet = wb.worksheets[index];
    if (!sheet) return null;
    if (sheet.name === name) return null;
    const invalid = SheetName.validationError(name);
    if (invalid) return invalid;
    const existing = wb.findSheet(name);
    if (existing && existing !== sheet) return `A worksheet named '${name}' already exists.`;
    this.execute(new RenameSheetCommand(sheet.name, name));
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
    const recalc = this.#engine.apply(EditResult.structuralChange);
    const circular = recalc.calculated ? recalc.hasCircularReference : this.#circular.peek();
    batch(() => {
      this.#circular.value = circular;
      this.#modified.value = true;
      this.#edits++;
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
