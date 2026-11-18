import { DateConversion } from '../../formulas/evaluation/dateConversion.ts';
import type { Expr } from '../../formulas/ast/expr.ts';
import { CellRange } from '../../model/cellRange.ts';
import { MergedRegion } from '../../model/mergedRegion.ts';
import { SheetDimensions } from '../../model/sheetDimensions.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import { XmlEventReader, type XmlAttribute, type XmlEventHandler, type XmlName } from '../../xml/index.ts';
import { excelToInternal, importFormula, shiftFormula } from './xlsxFormulas.ts';
import { attrIn, flag, RelType, relIdIn, type Relationship } from './xlsxPackage.ts';
import type { XlsxLeftOut, XlsxReportBuilder } from './xlsxReport.ts';
import { importListValidation, type XlsxValidation } from './xlsxValidation.ts';

/** Excel's 1904 date system counts days from 1 January 1904, 1462 days after the 1900 system's day 0. */
const DAYS_1904 = 1462;
/** Excel's default row height in points, when the sheet doesn't give one. */
const DEFAULT_ROW_POINTS = 15;
/** The most cells one hyperlink over a range is copied to. */
const HYPERLINK_CELL_LIMIT = 10_000;

/** What reading one worksheet needs from the rest of the workbook. */
export interface SheetReadContext {
  readonly sheet: Worksheet;
  /** The pooled style id for a cell format (`s` / `style`) of the file. */
  styleId(xf: number): number;
  /** Whether numbers in this cell format are dates that the 1904 date system shifts. */
  shiftsDate(xf: number): boolean;
  /** Whether numbers in this cell format are TRUE/FALSE. */
  isBoolean(xf: number): boolean;
  readonly sharedStrings: readonly string[];
  /** The worksheet part's relationships, for hyperlinks. */
  readonly rels: ReadonlyMap<string, Relationship>;
  hasName(name: string): boolean;
  /** The range a workbook or sheet name stands for, or null when it names none. */
  namedRange(name: string): CellRange | null;
  readonly report: XlsxReportBuilder;
}

/** `_x000D_`-style escapes OOXML uses for characters XML can't hold. */
export function decodeEscapes(text: string): string {
  if (!text.includes('_x')) return text;
  return text.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) =>
    String.fromCharCode(Number.parseInt(hex, 16)),
  );
}

/** `A1` (with optional `$`) → zero-based [row, column], or null when malformed or out of bounds. */
export function parseCellRef(ref: string): [row: number, column: number] | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(ref);
  if (!m) return null;
  const column = SheetDimensions.lettersToColumn(m[1]!);
  const row = Number(m[2]) - 1;
  if (column === null || !SheetDimensions.isValidRow(row)) return null;
  return [row, column];
}

/** An `xsd:dateTime` (the `t="d"` cell type) → a date serial, or null. */
function isoToSerial(text: string): number | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?/.exec(text.trim());
  if (!m) return null;
  return DateConversion.toSerial({
    year: Number(m[1]),
    month: Number(m[2]),
    day: Number(m[3]),
    hour: Number(m[4] ?? 0),
    minute: Number(m[5] ?? 0),
    second: Number(m[6] ?? 0),
    millisecond: Number((m[7] ?? '0').padEnd(3, '0')),
  });
}

/** Elements directly under `<worksheet>` that hold what Cellestra leaves out. */
const LEFT_OUT_PARTS: Readonly<Record<string, XlsxLeftOut>> = {
  conditionalFormatting: 'conditionalFormatting',
  autoFilter: 'autoFilters',
  drawing: 'drawings',
  picture: 'drawings',
  tableParts: 'tables',
};

interface SharedMaster {
  readonly expr: Expr;
  readonly row: number;
  readonly column: number;
}

/** The cell being read, from `<c>` to `</c>`. */
interface PendingCell {
  row: number;
  column: number;
  type: string;
  xf: number;
  value: string | null;
  inline: string;
  formula: string | null;
  formulaType: string;
  sharedIndex: string | null;
}

/** A `<dataValidation>` being read, with its formula and (in the `x14` form) range as text. */
type PendingValidation = { -readonly [K in keyof XlsxValidation]: XlsxValidation[K] };

/**
 * Streams one worksheet part into a sheet (PRD §8.19): cells with values, formulas and styles, row
 * heights and column widths, hidden lines, merges, frozen panes, hyperlinks and dropdown lists.
 */
class SheetReader implements XmlEventHandler {
  readonly #ctx: SheetReadContext;
  readonly #date1904: boolean;
  readonly #stack: string[] = [];
  readonly #masters = new Map<string, SharedMaster | null>();
  #row = -1;
  #column = -1;
  #cell: PendingCell | null = null;
  #validation: PendingValidation | null = null;
  /** Which text of `#validation` is being read, else null. */
  #validationText: 'formula1' | 'sqref' | null = null;
  /** Text of the `v`, `f` or inline `t` element being read, else null. */
  #text: string | null = null;
  #phonetic = 0;
  #defaultRowPoints = DEFAULT_ROW_POINTS;
  #sawView = false;

  constructor(ctx: SheetReadContext, date1904: boolean) {
    this.#ctx = ctx;
    this.#date1904 = date1904;
  }

  declaration(): void {}
  comment(): void {}
  processingInstruction(): void {}

  text(value: string): void {
    if (this.#text !== null && this.#phonetic === 0) this.#text += value;
    if (this.#validation && this.#validationText) this.#validation[this.#validationText] += value;
  }

  openElement(name: XmlName, attributes: XmlAttribute[]): void {
    const parent = this.#stack[this.#stack.length - 1];
    this.#stack.push(name.local);
    switch (name.local) {
      case 'c':
        if (parent === 'row') this.#openCell(attributes);
        return;
      case 'v':
      case 'f':
        if (parent === 'formula1' && this.#validation) this.#validationText = 'formula1';
        if (parent === 'c') {
          this.#text = '';
          if (name.local === 'f' && this.#cell) {
            this.#cell.formulaType = attrIn(attributes, 't') ?? 'normal';
            this.#cell.sharedIndex = attrIn(attributes, 'si');
          }
        }
        return;
      case 't':
        if (this.#cell && this.#stack.includes('is')) this.#text = '';
        return;
      case 'rPh':
        this.#phonetic++;
        return;
      case 'row':
        if (parent === 'sheetData') this.#openRow(attributes);
        return;
      case 'col':
        if (parent === 'cols') this.#columnLayout(attributes);
        return;
      case 'sheetFormatPr': {
        const points = Number(attrIn(attributes, 'defaultRowHeight'));
        if (points > 0) this.#defaultRowPoints = points;
        return;
      }
      case 'pane':
        if (parent === 'sheetView' && !this.#sawView) this.#pane(attributes);
        return;
      case 'mergeCell':
        this.#merge(attributes);
        return;
      case 'hyperlink':
        if (parent === 'hyperlinks') this.#hyperlink(attributes);
        return;
      case 'dataValidation':
        if (parent === 'dataValidations') this.#openValidation(attributes);
        return;
      case 'formula1':
        // The worksheet's own rules hold the formula as text; the x14 form wraps it in `<xm:f>`.
        if (parent === 'dataValidation' && this.#validation) this.#validationText = 'formula1';
        return;
      case 'sqref':
        if (parent === 'dataValidation' && this.#validation) this.#validationText = 'sqref';
        return;
      default: {
        const leftOut = parent === 'worksheet' ? LEFT_OUT_PARTS[name.local] : undefined;
        if (leftOut) this.#ctx.report.leftOut.add(leftOut);
      }
    }
  }

  closeElement(): void {
    const local = this.#stack.pop();
    switch (local) {
      case 'v':
        if (this.#cell && this.#text !== null) this.#cell.value = this.#text;
        this.#text = null;
        return;
      case 'f':
        if (this.#cell && this.#text !== null) this.#cell.formula = this.#text;
        this.#text = null;
        return;
      case 't':
        if (this.#cell && this.#text !== null) this.#cell.inline += this.#text;
        this.#text = null;
        return;
      case 'rPh':
        this.#phonetic--;
        return;
      case 'sheetView':
        this.#sawView = true;
        return;
      case 'formula1':
      case 'sqref':
        this.#validationText = null;
        return;
      case 'dataValidation':
        if (this.#validation) this.#finishValidation(this.#validation);
        this.#validation = null;
        return;
      case 'c':
        if (this.#cell) this.#finishCell(this.#cell);
        this.#cell = null;
        return;
      case undefined:
      default:
        return;
    }
  }

  #openRow(attributes: readonly XmlAttribute[]): void {
    const r = Number(attrIn(attributes, 'r'));
    this.#row = Number.isInteger(r) && r > 0 ? r - 1 : this.#row + 1;
    this.#column = -1;
    if (!SheetDimensions.isValidRow(this.#row)) return;
    const points = Number(attrIn(attributes, 'ht'));
    const custom = flag(attrIn(attributes, 'customHeight'));
    const hidden = flag(attrIn(attributes, 'hidden'));
    const styled = flag(attrIn(attributes, 'customFormat'));
    const xf = Number(attrIn(attributes, 's') ?? 0);
    const sized = points > 0 && (custom || Math.abs(points - this.#defaultRowPoints) > 0.01);
    if (!sized && !hidden && !(styled && xf > 0)) return;
    const row = this.#ctx.sheet.getOrCreateRow(this.#row);
    if (sized) row.height = Math.max(1, Math.round((points * 4) / 3));
    if (hidden) row.manuallyHidden = true;
    if (styled && xf > 0) row.styleId = this.#ctx.styleId(xf);
  }

  #columnLayout(attributes: readonly XmlAttribute[]): void {
    const min = Number(attrIn(attributes, 'min'));
    const max = Math.min(Number(attrIn(attributes, 'max')), SheetDimensions.maxColumns);
    if (!(min >= 1) || !(max >= min)) return;
    const width = Number(attrIn(attributes, 'width'));
    const hidden = flag(attrIn(attributes, 'hidden'));
    const xf = Number(attrIn(attributes, 'style') ?? 0);
    const styleId = xf > 0 ? this.#ctx.styleId(xf) : 0;
    for (let c = min - 1; c < max; c++) {
      const column = this.#ctx.sheet.getOrCreateColumn(c);
      // Excel stores widths in units of 1/7 of its 7-pixel default digit (8.43 characters → 9.14).
      if (width > 0) column.width = Math.max(1, Math.round(width * 7));
      if (hidden) column.hidden = true;
      if (styleId !== 0) column.styleId = styleId;
    }
  }

  #pane(attributes: readonly XmlAttribute[]): void {
    const state = attrIn(attributes, 'state');
    if (state !== 'frozen' && state !== 'frozenSplit') return;
    const columns = Number(attrIn(attributes, 'xSplit') ?? 0);
    const rows = Number(attrIn(attributes, 'ySplit') ?? 0);
    if (Number.isInteger(columns) && SheetDimensions.isValidColumn(columns))
      this.#ctx.sheet.frozenColumns = columns;
    if (Number.isInteger(rows) && SheetDimensions.isValidRow(rows)) this.#ctx.sheet.frozenRows = rows;
  }

  #merge(attributes: readonly XmlAttribute[]): void {
    const range = CellRange.tryParse(attrIn(attributes, 'ref'));
    if (!range || (range.rowCount === 1 && range.columnCount === 1)) return;
    const sheet = this.#ctx.sheet;
    // A file with overlapping merges is damaged; keep the first and drop the rest.
    if (sheet.merges.some((m) => m.range.intersects(range))) return;
    sheet.addMerge(new MergedRegion(range));
  }

  #hyperlink(attributes: readonly XmlAttribute[]): void {
    const range = CellRange.tryParse(attrIn(attributes, 'ref'));
    if (!range) return;
    const id = relIdIn(attributes);
    const rel = id === null ? undefined : this.#ctx.rels.get(id);
    const location = attrIn(attributes, 'location');
    let target: string | null = null;
    if (rel?.external && rel.type.endsWith(RelType.hyperlink))
      target = location ? `${rel.target}#${location}` : rel.target;
    else if (location) target = `#${excelToInternal(location) ?? location}`;
    if (target === null) return;
    const tooltip = attrIn(attributes, 'tooltip');
    const link = tooltip ? { target, tooltip } : { target };
    let count = 0;
    for (let r = range.top; r <= range.bottom; r++)
      for (let c = range.left; c <= range.right; c++) {
        if (count++ >= HYPERLINK_CELL_LIMIT) return;
        this.#ctx.sheet.getOrCreate(r, c).hyperlink = link;
      }
  }

  #openValidation(attributes: readonly XmlAttribute[]): void {
    this.#validation = {
      type: attrIn(attributes, 'type'),
      sqref: attrIn(attributes, 'sqref') ?? '',
      formula1: '',
      errorStyle: attrIn(attributes, 'errorStyle'),
      showErrorMessage: flag(attrIn(attributes, 'showErrorMessage')),
      showInputMessage: flag(attrIn(attributes, 'showInputMessage')),
      error: attrIn(attributes, 'error'),
      prompt: attrIn(attributes, 'prompt'),
    };
  }

  #finishValidation(rule: PendingValidation): void {
    const rules = importListValidation(rule, (name) => this.#ctx.namedRange(name));
    if (rules === null) this.#ctx.report.leftOut.add('dataValidation');
    else for (const v of rules) this.#ctx.sheet.addValidation(v);
  }

  #openCell(attributes: readonly XmlAttribute[]): void {
    const ref = attrIn(attributes, 'r');
    const at = ref === null ? null : parseCellRef(ref);
    const row = at ? at[0] : this.#row;
    const column = at ? at[1] : this.#column + 1;
    this.#column = column;
    if (!SheetDimensions.isValidRow(row) || !SheetDimensions.isValidColumn(column)) return;
    this.#cell = {
      row,
      column,
      type: attrIn(attributes, 't') ?? 'n',
      xf: Number(attrIn(attributes, 's') ?? 0) || 0,
      value: null,
      inline: '',
      formula: null,
      formulaType: 'normal',
      sharedIndex: null,
    };
  }

  #cachedValue(cell: PendingCell): CellValue {
    const v = cell.value;
    switch (cell.type) {
      case 's': {
        const text = v === null ? undefined : this.#ctx.sharedStrings[Number(v)];
        return text === undefined ? CellValue.empty : CellValue.text(text);
      }
      case 'inlineStr':
        return CellValue.text(decodeEscapes(cell.inline));
      case 'str':
        return v === null ? CellValue.empty : CellValue.text(decodeEscapes(v));
      case 'b':
        return v === null ? CellValue.empty : CellValue.boolean(v.trim() === '1' || v.trim() === 'true');
      case 'e': {
        if (v === null) return CellValue.empty;
        const code = CellError.codeFromToken(v.trim());
        return code === null
          ? CellValue.text(v)
          : CellValue.error(new CellError(code, `The file saved this result as ${v.trim()}.`));
      }
      case 'd': {
        const serial = v === null ? null : isoToSerial(v);
        return serial === null ? CellValue.empty : CellValue.number(serial);
      }
      default: {
        if (v === null || v.trim() === '') return CellValue.empty;
        const n = Number(v);
        if (!Number.isFinite(n)) return CellValue.empty;
        if (this.#ctx.isBoolean(cell.xf)) return CellValue.boolean(n !== 0);
        return CellValue.number(this.#date1904 && this.#ctx.shiftsDate(cell.xf) ? n + DAYS_1904 : n);
      }
    }
  }

  /** The cell's formula as Cellestra text, or null when it is kept as its value (and counted). */
  #formula(cell: PendingCell): string | null {
    if (cell.formula === null) return null;
    const text = cell.formula.trim();
    if (cell.formulaType === 'array' || cell.formulaType === 'dataTable') {
      this.#ctx.report.formulasAsValues++;
      return null;
    }
    if (cell.formulaType === 'shared' && cell.sharedIndex !== null) {
      if (text === '') {
        const master = this.#masters.get(cell.sharedIndex);
        if (master) return shiftFormula(master.expr, cell.row - master.row, cell.column - master.column);
        this.#ctx.report.formulasAsValues++;
        return null;
      }
      const imported = importFormula(text, (n) => this.#ctx.hasName(n));
      this.#masters.set(
        cell.sharedIndex,
        imported.ok ? { expr: imported.expr, row: cell.row, column: cell.column } : null,
      );
      if (imported.ok) return imported.formula;
      this.#ctx.report.formulasAsValues++;
      return null;
    }
    if (text === '') return null;
    const imported = importFormula(text, (n) => this.#ctx.hasName(n));
    if (imported.ok) return imported.formula;
    this.#ctx.report.formulasAsValues++;
    return null;
  }

  #finishCell(cell: PendingCell): void {
    const value = this.#cachedValue(cell);
    const formula = this.#formula(cell);
    const styleId = cell.xf > 0 ? this.#ctx.styleId(cell.xf) : 0;
    if (formula === null && value.kind === 'empty' && styleId === 0) return;
    const target = this.#ctx.sheet.getOrCreate(cell.row, cell.column);
    target.formula = formula;
    target.value = value;
    target.styleId = styleId;
  }
}

/** Read one worksheet part, given as text pieces, into `ctx.sheet`. */
export function readWorksheet(chunks: Iterable<string>, ctx: SheetReadContext, date1904: boolean): void {
  const reader = new XmlEventReader(new SheetReader(ctx, date1904));
  for (const chunk of chunks) reader.write(chunk);
  reader.close();
}
