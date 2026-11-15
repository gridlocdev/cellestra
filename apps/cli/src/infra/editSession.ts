import { CellRef, writeCsv, type EditCommand, type Workbook } from '@cellestra/core';
import { extname } from 'node:path';
import { CliError } from './errors.ts';
import type { Console } from './io.ts';
import { address, cellReportAt, errorReports, writeReport } from './report.ts';
import {
  load,
  recalculate,
  recalculateAfterEdits,
  saveOds,
  saveText,
  type LoadedWorkbook,
  type WorkbookFormat,
} from './workspace.ts';

const JSON_TO_STDOUT =
  "--json writes the report to stdout, so the workbook can't go there too — give -o <file> or --in-place.";

/**
 * One load → edit → recalculate (as the workbook's calculation mode says) → atomic save, with the report every writing verb prints: the
 * cells it changed, every error cell now in the workbook, and any notes and extra fields.
 */
export class EditSession {
  readonly loaded: LoadedWorkbook;
  readonly locale: string;
  readonly con: Console;
  readonly #changed: CellRef[] = [];
  readonly #notes: string[] = [];
  readonly #fields = new Map<string, unknown>();
  #recalculateAll = false;

  private constructor(con: Console, loaded: LoadedWorkbook, locale: string) {
    this.con = con;
    this.loaded = loaded;
    this.locale = locale;
  }

  get workbook(): Workbook {
    return this.loaded.workbook;
  }

  apply(edit: EditCommand): void {
    const result = edit.apply(this.workbook);
    // A loop, not push(...cells): a whole-column edit changes more cells than the stack holds arguments.
    for (const c of result.changedCells) this.#changed.push(CellRef.of(c.sheet, c.row, c.column));
  }

  touched(sheet: string, row: number, column: number): void {
    this.#changed.push(CellRef.of(sheet, row, column));
  }

  /** Recompute every formula before saving, even in manual calculation (the `recalc` command). */
  recalculateAll(): void {
    this.#recalculateAll = true;
  }

  note(message: string): void {
    this.#notes.push(message);
  }

  field(name: string, value: unknown): void {
    this.#fields.set(name, value);
  }

  static run(
    con: Console,
    path: string,
    target: string,
    locale: string | null,
    json: boolean,
    edit: (session: EditSession) => void,
    format: WorkbookFormat | null = 'ods',
    dryRun = false,
  ): number {
    if (json && target === '-' && !dryRun) throw new CliError(JSON_TO_STDOUT);
    const loaded = load(con, path, format, locale);
    const session = new EditSession(con, loaded, locale ?? loaded.workbook.settings.effectiveLocale);
    edit(session);
    return session.#finish(target, json, dryRun);
  }

  static runLoaded(
    con: Console,
    loaded: LoadedWorkbook,
    target: string,
    locale: string | null,
    json: boolean,
    edit: (session: EditSession) => void,
  ): number {
    if (json && target === '-') throw new CliError(JSON_TO_STDOUT);
    const session = new EditSession(con, loaded, locale ?? loaded.workbook.settings.effectiveLocale);
    edit(session);
    return session.#finish(target, json, false);
  }

  /** A .csv or .tsv target gets the active sheet as text; anything else (or stdout) gets .ods. */
  #save(target: string): void {
    const ext = target === '-' ? '' : extname(target).toLowerCase();
    if (ext !== '.csv' && ext !== '.tsv') {
      saveOds(this.con, this.workbook, this.loaded.package, target);
      return;
    }
    const sheet = this.workbook.activeSheet;
    const others = this.workbook.sheetCount - 1;
    if (others > 0) this.note(`${ext.slice(1).toUpperCase()} keeps one sheet: wrote '${sheet.name}' only.`);
    const delimiter = ext === '.tsv' ? '\t' : ',';
    saveText(this.con, writeCsv(this.workbook, sheet, { delimiter, locale: this.locale }), target);
  }

  #finish(target: string, json: boolean, dryRun: boolean): number {
    const result = this.#recalculateAll ? recalculate(this.workbook) : recalculateAfterEdits(this.workbook);
    if (!dryRun) this.#save(target);
    if (!json) {
      for (const note of this.#notes) this.con.errLine(note);
      if (dryRun) this.con.errLine('Dry run: nothing written.');
      if (result.hasCircularReference) {
        this.con.errLine(`Warning: ${result.circularCells.length} circular reference(s).`);
      }
      return 0;
    }
    // Distinct, in first-touched order; a cell whose sheet was deleted has nothing to report.
    const seen = new Set<string>();
    const changed = [];
    for (const c of this.#changed) {
      const key = `${c.sheet}\u0000${c.row}\u0000${c.column}`;
      if (seen.has(key)) continue;
      seen.add(key);
      if (this.workbook.findSheet(c.sheet)) changed.push(cellReportAt(this.workbook, c, this.locale));
    }
    const payload: Record<string, unknown> = {
      ok: true,
      file: target,
      dryRun: dryRun ? true : null,
      changed,
      recalculated: result.evaluated,
      errors: errorReports(this.workbook, this.locale),
      circular: result.circularCells.map((c) => address(c.sheet, c.row, c.column)),
    };
    if (this.#notes.length > 0) payload['notes'] = this.#notes;
    for (const [key, value] of this.#fields) payload[key] = value;
    writeReport(this.con, payload);
    return 0;
  }
}
