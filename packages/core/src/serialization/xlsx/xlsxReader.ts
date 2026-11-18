import { Defaults } from '../../defaults.ts';
import { CellRange } from '../../model/cellRange.ts';
import { SheetName } from '../../model/sheetName.ts';
import { Workbook } from '../../model/workbook.ts';
import { NumberFormats } from '../../styling/cellStyle.ts';
import { XmlEventReader, type XmlElement, type XmlEventHandler, type XmlName } from '../../xml/index.ts';
import { OdsUnsupportedFileError } from '../ods/odsPackage.ts';
import { readPalette } from './xlsxColors.ts';
import { excelToInternal } from './xlsxFormulas.ts';
import { at, flag, kid, kids, RelType, relIdIn, XlsxPackage, type Relationship } from './xlsxPackage.ts';
import { XlsxReportBuilder, type XlsxImportReport } from './xlsxReport.ts';
import { decodeEscapes, readWorksheet } from './xlsxSheetRead.ts';
import { readStyles } from './xlsxStyles.ts';

/** A workbook read from an `.xlsx`, and what the import left out (PRD §8.19). */
export interface XlsxReadResult {
  readonly workbook: Workbook;
  readonly report: XlsxImportReport;
}

/** Collects the `<si>` texts of `sharedStrings.xml` as they stream past; phonetic runs (`rPh`) are skipped. */
class SharedStrings implements XmlEventHandler {
  readonly strings: string[] = [];
  readonly #stack: string[] = [];
  #current = '';

  declaration(): void {}
  comment(): void {}
  processingInstruction(): void {}

  openElement(name: XmlName): void {
    this.#stack.push(name.local);
    if (name.local === 'si') this.#current = '';
  }

  closeElement(): void {
    const local = this.#stack.pop();
    if (local === 'si') this.strings.push(decodeEscapes(this.#current));
  }

  text(value: string): void {
    const n = this.#stack.length;
    if (this.#stack[n - 1] === 't' && !this.#stack.includes('rPh')) this.#current += value;
  }
}

function readSharedStrings(chunks: Iterable<string> | null): string[] {
  if (chunks === null) return [];
  const handler = new SharedStrings();
  const reader = new XmlEventReader(handler);
  for (const chunk of chunks) reader.write(chunk);
  reader.close();
  return handler.strings;
}

/** A sheet name Cellestra accepts and no other sheet has, made from the file's. */
function usableSheetName(workbook: Workbook, raw: string): string {
  let base = raw.replace(/[:\\/?*[\]]/g, '_').replace(/^'+|'+$/g, '');
  if (base.trim() === '') base = 'Sheet';
  base = base.slice(0, SheetName.maxLength);
  if (SheetName.validationError(base) !== null) base = 'Sheet';
  const taken = (n: string): boolean =>
    workbook.worksheets.some((s) => s.name.toLowerCase() === n.toLowerCase());
  if (!taken(base)) return base;
  for (let n = 2; ; n++) {
    const suffix = ` ${n}`;
    const candidate = base.slice(0, SheetName.maxLength - suffix.length) + suffix;
    if (!taken(candidate)) return candidate;
  }
}

/** The first part of a type in a part's relationships, or null. */
function relTarget(rels: ReadonlyMap<string, Relationship>, type: string): string | null {
  for (const rel of rels.values()) if (!rel.external && rel.type.endsWith(type)) return rel.target;
  return null;
}

interface SheetEntry {
  readonly name: string;
  readonly path: string;
}

/**
 * Reads `.xlsx` bytes into a new workbook (PRD §8.19). Throws `OdsUnsupportedFileError` with a
 * message for the person for an encrypted workbook, an `.xls`, or bytes that aren't an `.xlsx`.
 */
export const XlsxReader = {
  read(bytes: Uint8Array, locale = 'en-US'): XlsxReadResult {
    const pkg = XlsxPackage.read(bytes);
    const report = new XlsxReportBuilder();
    const workbookPath = pkg.workbookPath();
    const book = pkg.xml(workbookPath);
    if (!book) throw new OdsUnsupportedFileError('This file is damaged or is not an .xlsx workbook.');
    const rels = pkg.rels(workbookPath);

    const stylesPath = relTarget(rels, RelType.styles);
    const themePath = relTarget(rels, RelType.theme);
    const stylesXml = stylesPath ? pkg.xml(stylesPath) : null;
    const palette = readPalette(themePath ? pkg.xml(themePath) : null, stylesXml);
    const { cellStyles, booleanFormats } = readStyles(stylesXml, palette, locale);
    const stringsPath = relTarget(rels, RelType.sharedStrings);
    const sharedStrings = readSharedStrings(stringsPath ? pkg.chunks(stringsPath) : null);

    const workbook = new Workbook();
    const date1904 = flag(at(kid(book, 'workbookPr'), 'date1904'));
    if (at(kid(book, 'calcPr'), 'calcMode') === 'manual') workbook.settings.calculationMode = 'manual';

    // Sheets first, so names and formulas can see them all.
    const entries: SheetEntry[] = [];
    const fileIndexToSheet = new Map<number, string>();
    kids(kid(book, 'sheets'), 'sheet').forEach((el, index) => {
      const id = relIdIn(el.attributes);
      const rel = id === null ? undefined : rels.get(id);
      if (!rel || rel.external) return;
      if (rel.type.endsWith(RelType.chartsheet)) return void report.leftOut.add('chartSheets');
      if (!rel.type.endsWith(RelType.worksheet)) return void report.leftOut.add('otherSheets');
      const state = at(el, 'state');
      if (state === 'hidden' || state === 'veryHidden') report.hiddenSheetsShown++;
      const name = usableSheetName(workbook, at(el, 'name') ?? '');
      workbook.addSheetForLoad(name);
      fileIndexToSheet.set(index, name);
      entries.push({ name, path: rel.target });
    });
    if (entries.length === 0) workbook.addSheetForLoad(Defaults.firstSheetName);

    const activeTab = Number(at(kid(kid(book, 'bookViews'), 'workbookView'), 'activeTab') ?? 0);
    const activeName = fileIndexToSheet.get(activeTab);
    const activeIndex = workbook.worksheets.findIndex((s) => s.name === activeName);
    workbook.settings.activeSheetIndex = Math.max(0, activeIndex);

    readDefinedNames(kid(book, 'definedNames'), workbook, fileIndexToSheet, report);

    const styleIds = new Map<number, number>();
    const styleId = (xf: number): number => {
      let id = styleIds.get(xf);
      if (id === undefined) {
        const style = cellStyles[xf];
        id = style ? workbook.styles.intern(style) : 0;
        styleIds.set(xf, id);
      }
      return id;
    };
    const shiftsDate = (xf: number): boolean => {
      const style = cellStyles[xf];
      if (!style) return false;
      const kind = NumberFormats.kind(style.numberFormat);
      return kind === 'date' || kind === 'dateTime';
    };

    for (const entry of entries) {
      const sheet = workbook.worksheets.find((s) => s.name === entry.name)!;
      const sheetRels = pkg.rels(entry.path);
      for (const rel of sheetRels.values()) {
        if (rel.type.endsWith(RelType.comments)) report.leftOut.add('comments');
        else if (rel.type.endsWith(RelType.pivotTable)) report.leftOut.add('pivotTables');
      }
      const chunks = pkg.chunks(entry.path);
      if (chunks === null) continue;
      readWorksheet(
        chunks,
        {
          sheet,
          styleId,
          shiftsDate,
          isBoolean: (xf) => booleanFormats.has(xf),
          sharedStrings,
          rels: sheetRels,
          hasName: (n) => workbook.findNamedRange(n, sheet.name) !== undefined,
          report,
        },
        date1904,
      );
    }

    if (relTarget(rels, RelType.vbaProject) !== null) report.leftOut.add('macros');
    if (kid(book, 'pivotCaches')) report.leftOut.add('pivotTables');
    return { workbook, report: report.build() };
  },
};

/**
 * `definedNames`: each name that is one range becomes a named range, scoped to its sheet when it
 * has `localSheetId`. Excel's own names (`_xlnm.Print_Area`, filter ranges) and hidden names are
 * skipped silently; a constant or formula name is left out and reported.
 */
function readDefinedNames(
  container: XmlElement | null,
  workbook: Workbook,
  fileIndexToSheet: ReadonlyMap<number, string>,
  report: XlsxReportBuilder,
): void {
  for (const el of kids(container, 'definedName')) {
    const name = at(el, 'name');
    if (name === null || name.startsWith('_xlnm.') || flag(at(el, 'hidden'))) continue;
    const local = at(el, 'localSheetId');
    const scope = local === null ? null : (fileIndexToSheet.get(Number(local)) ?? undefined);
    if (scope === undefined) continue;
    const internal = excelToInternal(el.value.trim());
    const range = internal === null ? null : CellRange.tryParse(internal);
    if (!range || range.sheet === null || (range.end.sheet !== null && range.end.sheet !== range.sheet)) {
      report.leftOut.add('names');
      continue;
    }
    const key = name.toLowerCase();
    const taken = workbook.namedRanges.some(
      (n) => n.name.toLowerCase() === key && (n.scope ?? null)?.toLowerCase() === scope?.toLowerCase(),
    );
    if (taken) continue;
    workbook.addNamedRange(scope === null ? { name, target: range } : { name, target: range, scope });
  }
}
