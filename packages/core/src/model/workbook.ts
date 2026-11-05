import type { ChartDefinition } from '@cellestra/charts';
import { Defaults, nextChartSheetName, nextSheetName } from '../defaults.ts';
import { StylePool } from '../styling/stylePool.ts';
import { DocumentSettings } from './documentSettings.ts';
import type { NamedRange } from './namedRange.ts';
import { SheetName } from './sheetName.ts';
import { Worksheet } from './worksheet.ts';

const same = (a: string | null | undefined, b: string | null | undefined): boolean =>
  (a ?? null) === null ? (b ?? null) === null : (b ?? null) !== null && a!.toLowerCase() === b!.toLowerCase();

/** The whole document: worksheets, named ranges, the style pool, and document settings. */
export class Workbook {
  private sheets: Worksheet[] = [];
  private names: NamedRange[] = [];
  readonly styles = new StylePool();
  readonly settings = new DocumentSettings();

  get worksheets(): readonly Worksheet[] {
    return this.sheets;
  }
  get namedRanges(): readonly NamedRange[] {
    return this.names;
  }
  get sheetCount(): number {
    return this.sheets.length;
  }

  get activeSheet(): Worksheet {
    if (this.sheets.length === 0) throw new Error('Workbook has no worksheets.');
    const i = Math.min(Math.max(this.settings.activeSheetIndex, 0), this.sheets.length - 1);
    return this.sheets[i]!;
  }

  /** A new workbook: one sheet named `Sheet1` (PRD §8.8). */
  static createNew(): Workbook {
    const wb = new Workbook();
    wb.addSheet(Defaults.firstSheetName);
    wb.settings.activeSheetIndex = 0;
    return wb;
  }

  /** Add a sheet; with no name, the next free `SheetN`. */
  addSheet(name?: string): Worksheet {
    const n = name ?? nextSheetName(this.sheets.map((s) => s.name));
    SheetName.validate(n);
    if (this.findSheet(n)) throw new Error(`A worksheet named '${n}' already exists.`);
    const sheet = new Worksheet(n);
    this.sheets.push(sheet);
    return sheet;
  }

  addChartSheet(chart: ChartDefinition): Worksheet {
    const sheet = this.addSheet(nextChartSheetName(this.sheets.map((s) => s.name)));
    sheet.chart = chart;
    return sheet;
  }

  /** Readers use this: files may carry names our rules would reject; uniqueness is still enforced. */
  addSheetForLoad(name: string): Worksheet {
    if (this.findSheet(name)) throw new Error(`A worksheet named '${name}' already exists.`);
    const sheet = new Worksheet(name);
    this.sheets.push(sheet);
    return sheet;
  }

  renameSheet(sheet: Worksheet, newName: string): void {
    SheetName.validate(newName);
    const existing = this.findSheet(newName);
    if (existing && existing !== sheet) throw new Error(`A worksheet named '${newName}' already exists.`);
    sheet.name = newName;
  }

  duplicateSheet(source: Worksheet): Worksheet {
    const index = this.sheets.indexOf(source);
    if (index < 0) throw new Error('The sheet to duplicate is not in this workbook.');
    const copy = source.clone(this.uniqueDuplicateName(source.name));
    // Tables persist as document-global ODF database ranges, so the copies need fresh names.
    const taken = new Set(this.sheets.flatMap((s) => s.tables).map((t) => t.name.toLowerCase()));
    for (const t of [...copy.tables]) {
      if (taken.has(t.name.toLowerCase())) {
        let n = 2;
        while (taken.has(`${t.name}_${n}`.toLowerCase())) n++;
        const fresh = `${t.name}_${n}`;
        copy.replaceTable(t, { ...t, name: fresh });
        taken.add(fresh.toLowerCase());
      } else taken.add(t.name.toLowerCase());
    }
    this.sheets.splice(index + 1, 0, copy);
    return copy;
  }

  private uniqueDuplicateName(base: string): string {
    for (let n = 0; ; n++) {
      const suffix = n === 0 ? ' copy' : ` copy ${n + 1}`;
      const headMax = Math.max(1, SheetName.maxLength - suffix.length);
      const candidate = (base.length > headMax ? base.slice(0, headMax) : base) + suffix;
      if (!this.findSheet(candidate)) return candidate;
    }
  }

  insertSheet(index: number, sheet: Worksheet): void {
    SheetName.validate(sheet.name);
    if (this.findSheet(sheet.name)) throw new Error(`A worksheet named '${sheet.name}' already exists.`);
    this.sheets.splice(Math.min(Math.max(index, 0), this.sheets.length), 0, sheet);
  }

  removeSheet(sheet: Worksheet): boolean {
    if (this.sheets.length <= 1) throw new Error('A workbook must keep at least one worksheet.');
    const i = this.sheets.indexOf(sheet);
    if (i < 0) return false;
    this.sheets.splice(i, 1);
    if (this.settings.activeSheetIndex >= this.sheets.length)
      this.settings.activeSheetIndex = this.sheets.length - 1;
    return true;
  }

  moveSheet(from: number, to: number): void {
    if (from < 0 || from >= this.sheets.length) throw new RangeError('fromIndex out of range');
    const target = Math.min(Math.max(to, 0), this.sheets.length - 1);
    const [sheet] = this.sheets.splice(from, 1);
    this.sheets.splice(target, 0, sheet!);
  }

  /** Case-insensitive. */
  findSheet(name: string): Worksheet | undefined {
    const lower = name.toLowerCase();
    return this.sheets.find((s) => s.name.toLowerCase() === lower);
  }

  indexOf(sheet: Worksheet): number {
    return this.sheets.indexOf(sheet);
  }

  // ---- Named ranges -----------------------------------------------------

  /** A sheet-local name shadows a global one of the same name. */
  findNamedRange(name: string, sheetScope: string | null = null): NamedRange | undefined {
    if (sheetScope !== null) {
      const local = this.names.find((n) => same(n.scope, sheetScope) && same(n.name, name));
      if (local) return local;
    }
    return this.names.find((n) => (n.scope ?? null) === null && same(n.name, name));
  }

  addNamedRange(nr: NamedRange): void {
    if (this.names.some((n) => same(n.name, nr.name) && same(n.scope, nr.scope)))
      throw new Error(`A named range '${nr.name}' already exists in this scope.`);
    this.names.push(nr);
  }

  removeNamedRange(name: string, scope: string | null = null): boolean {
    const i = this.names.findIndex((n) => same(n.name, name) && same(n.scope, scope));
    if (i < 0) return false;
    this.names.splice(i, 1);
    return true;
  }
}
