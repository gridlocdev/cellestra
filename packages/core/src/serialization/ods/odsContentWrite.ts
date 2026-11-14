import { Defaults } from '../../defaults.ts';
import { DateConversion, type CivilDateTime } from '../../formulas/evaluation/dateConversion.ts';
import { OpenFormulaTranslator } from '../../formulas/openFormulaTranslator.ts';
import type { Cell } from '../../model/cell.ts';
import type { MergedRegion } from '../../model/mergedRegion.ts';
import type { CellRange } from '../../model/cellRange.ts';
import type { NamedRange } from '../../model/namedRange.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { cellKey } from '../../model/sheetDimensions.ts';
import { NumberFormats } from '../../styling/cellStyle.ts';
import { NumberFormatter } from '../../styling/numberFormatter.ts';
import { StylePool } from '../../styling/stylePool.ts';
import { roundTripNumber } from '../../values/cellValue.ts';
import { parseXmlElement, XmlDocument, XmlElement, XmlText } from '../../xml/index.ts';
import { OdsCharts } from './odsCharts.ts';
import { checkboxCells, checkboxControl, checkboxForm, Form, formsElement, Ooo } from './odsCheckboxes.ts';
import { appendParagraphText, odfCellAddress, odfRangeAddress } from './odsContentShared.ts';
import { addDatabaseRanges } from './odsDatabaseRanges.ts';
import type { ForeignXml } from './odsPackage.ts';
import { OdsPreserved } from './odsPreserved.ts';
import { NO_PAGE_PLAN, isGeneratedTableStyle, type PagePlan } from './odsPageSetup.ts';
import { StyleRegistry } from './odsStyleRegistry.ts';
import { addContentValidations } from './odsValidation.ts';
import {
  Calcext,
  Dc,
  Draw,
  Fo,
  NumberNs,
  OdsXmlNames as N,
  Of,
  Office,
  Style,
  Svg,
  Table,
  Text,
  Xlink,
} from './odsXmlNames.ts';

/** A fixed locale for the cached display text written beside each value, whatever the system's. */
const INVARIANT_LOCALE = 'en-US';

type ValidationTargets = ReadonlyArray<{ readonly range: CellRange; readonly name: string }>;
const NO_VALIDATIONS: ValidationTargets = [];

/**
 * Build `content.xml`: the cell grid, automatic styles, merges, column widths/row heights, hidden
 * lines, named ranges, validations, and database ranges. Unmodeled `office:spreadsheet` children
 * captured on read are re-emitted verbatim in their original position relative to the tables, and
 * so is the table, cell and named-expression content kept by `OdsPreserved`.
 */
export function writeContentXml(
  workbook: Workbook,
  foreign: readonly ForeignXml[],
  reservedStyleNames: ReadonlySet<string> = new Set(),
  commonStyles: XmlElement[] | null = null,
  pages: PagePlan = NO_PAGE_PLAN,
): XmlDocument {
  const autoStyles = new XmlElement(Office('automatic-styles'));
  const preserved = OdsPreserved.workbook(workbook);
  // Generated names must not reuse a preserved automatic style's name, nor any styles.xml name.
  const reserved = new Set<string>(reservedStyleNames);
  for (const style of preserved?.automaticStyles ?? []) {
    // Table styles an earlier save wrote for page setups are written again below, renamed.
    if (isGeneratedTableStyle(style)) continue;
    const name = style.attr(Style('name'));
    if (name !== null) reserved.add(name);
    autoStyles.add(style.clone());
  }
  for (const style of pages.tableStyles) autoStyles.add(style.clone());
  const registry = new StyleRegistry(workbook.styles, autoStyles, reserved, reservedStyleNames, commonStyles);

  const spreadsheet = new XmlElement(Office('spreadsheet'));
  // Structure protection and its key live here; the model has no field for them.
  for (const a of preserved?.spreadsheetAttributes ?? []) spreadsheet.setAttr(a.name, a.value);

  for (const child of foreign) if (child.beforeTables) spreadsheet.add(parseXmlElement(child.xml));

  // Data validations are declared once, before the tables (ODF schema order), and referenced by
  // name from each validated cell.
  const validationNames = addContentValidations(spreadsheet, workbook);

  workbook.worksheets.forEach((sheet, index) => {
    spreadsheet.add(
      writeTable(workbook, sheet, index, registry, validationNames.get(sheet.name) ?? NO_VALIDATIONS, pages),
    );
  });

  addNamedExpressions(spreadsheet, workbook, preserved?.namedExpressions ?? []);
  addDatabaseRanges(spreadsheet, workbook);

  for (const child of foreign) if (!child.beforeTables) spreadsheet.add(parseXmlElement(child.xml));

  const root = new XmlElement(Office('document-content'))
    .declare(Office.prefix, Office.uri)
    .declare(Table.prefix, Table.uri)
    .declare(Text.prefix, Text.uri)
    .declare(Style.prefix, Style.uri)
    .declare(Fo.prefix, Fo.uri)
    .declare(NumberNs.prefix, NumberNs.uri)
    .declare(Svg.prefix, Svg.uri)
    .declare(Draw.prefix, Draw.uri)
    .declare(Form.prefix, Form.uri)
    .declare(Ooo.prefix, Ooo.uri)
    .declare(Xlink.prefix, Xlink.uri)
    .declare(Of.prefix, Of.uri)
    .declare(Calcext.prefix, Calcext.uri)
    // Comments (office:annotation) carry dc:creator and dc:date; declared once here, not per note.
    .declare(Dc.prefix, Dc.uri)
    .setAttr(Office('version'), '1.3')
    // Schema order: office:font-face-decls, office:automatic-styles, office:body.
    .add(preserved?.fontFaces?.clone(), autoStyles, new XmlElement(Office('body'), spreadsheet));

  return new XmlDocument(root);
}

interface SheetWriter {
  readonly sheet: Worksheet;
  readonly registry: StyleRegistry;
  readonly validations: ValidationTargets;
  readonly lastColumn: number;
  /** Checkbox cells (by `cellKey`) and the id of the control each gets. */
  readonly checkboxes: ReadonlyMap<number, string>;
  /** The merges covering each row, asked for in increasing row order. */
  readonly merges: MergeSweep;
}

/**
 * The merges that cover a row, for rows visited top to bottom: merges start being reported at
 * their top row and stop after their bottom one, so each row costs the merges open there, not a
 * scan of every merge in the sheet.
 */
class MergeSweep {
  private readonly sorted: readonly MergedRegion[];
  private next = 0;
  private open: MergedRegion[] = [];

  constructor(merges: readonly MergedRegion[]) {
    this.sorted = [...merges].sort((a, b) => a.range.top - b.range.top);
  }

  at(row: number): readonly MergedRegion[] {
    while (this.next < this.sorted.length && this.sorted[this.next]!.range.top <= row) {
      this.open.push(this.sorted[this.next++]!);
    }
    if (this.open.some((m) => m.range.bottom < row))
      this.open = this.open.filter((m) => m.range.bottom >= row);
    return this.open;
  }
}

/** What may come before `office:forms` in a table (ODF 1.3 §9.1.2); everything else comes after. */
const BEFORE_FORMS = ['title', 'desc', 'table-source', 'scenario'];

function writeTable(
  workbook: Workbook,
  sheet: Worksheet,
  sheetIndex: number,
  registry: StyleRegistry,
  validations: ValidationTargets,
  pages: PagePlan,
): XmlElement {
  const table = new XmlElement(Table('table')).setAttr(Table('name'), sheet.name);
  const preserved = OdsPreserved.table(sheet);
  for (const a of preserved?.attributes ?? []) table.setAttr(a.name, a.value);
  const pageStyle = pages.tableStyleNames.get(sheet);
  if (pageStyle !== undefined) table.setAttr(Table('style-name'), pageStyle);

  // Checkbox cells save as form controls: one form for the sheet's checkboxes, joined to the
  // file's own `office:forms` when it has one so its other controls are kept.
  const checkboxes = checkboxCells(workbook, sheet, sheetIndex);
  let form = checkboxes.size > 0 ? checkboxForm(sheet, checkboxes) : null;
  for (const child of preserved?.before ?? []) {
    const copy = child.clone();
    if (form && copy.is(Office('forms'))) {
      copy.add(form);
      form = null;
    } else if (form && !(copy.uri === Table.uri && BEFORE_FORMS.includes(copy.local))) {
      table.add(formsElement(form));
      form = null;
    }
    table.add(copy);
  }
  if (form) table.add(formsElement(form));

  const used = sheet.getUsedRange();
  let lastColumn = used?.right ?? 0;
  let lastRow = used?.bottom ?? 0;

  // A validated cell must be emitted (with its content-validation-name) even when it is blank and
  // lies past the used range, so extend the written bounds to cover every validation.
  for (const validation of sheet.validations) {
    lastColumn = Math.max(lastColumn, validation.target.right);
    lastRow = Math.max(lastRow, validation.target.bottom);
  }

  // A column that carries formatting but no cells is past the used range and still has to be
  // written, or the formatting is lost on save.
  for (const [index, model] of sheet.columns) if (!model.isDefault) lastColumn = Math.max(lastColumn, index);

  writeColumns(table, sheet, lastColumn, registry);
  const merges = new MergeSweep(sheet.merges);
  writeRows(table, { sheet, registry, validations, lastColumn, checkboxes, merges }, lastRow);

  // A chart sheet embeds its chart as an ODF draw frame (the object was written by writeObjects).
  if (sheet.isChartSheet) OdsCharts.injectFrame(table, sheet);

  addLocalNamedExpressions(table, workbook, sheet, preserved?.namedExpressions ?? []);
  for (const child of preserved?.after ?? []) table.add(child.clone());
  return table;
}

function writeColumns(
  table: XmlElement,
  sheet: Worksheet,
  lastColumn: number,
  registry: StyleRegistry,
): void {
  // One element per column up to the used width, coalescing identical neighbours.
  let column = 0;
  while (column <= lastColumn) {
    const signature = columnSignature(sheet, column, registry);
    let repeat = 1;
    while (
      column + repeat <= lastColumn &&
      sameSignature(columnSignature(sheet, column + repeat, registry), signature)
    ) {
      repeat++;
    }

    const element = new XmlElement(Table('table-column'));
    if (signature.style !== null) element.setAttr(Table('style-name'), signature.style);
    if (repeat > 1) element.setAttr(Table('number-columns-repeated'), repeat);
    if (signature.hidden) element.setAttr(Table('visibility'), 'collapse');
    // How a whole formatted column travels: one attribute, not a style on every cell in it.
    if (signature.defaultCellStyle !== null)
      element.setAttr(Table('default-cell-style-name'), signature.defaultCellStyle);

    table.add(element);
    column += repeat;
  }
}

interface ColumnSignature {
  readonly style: string | null;
  readonly hidden: boolean;
  readonly defaultCellStyle: string | null;
}

function sameSignature(a: ColumnSignature, b: ColumnSignature): boolean {
  return a.style === b.style && a.hidden === b.hidden && a.defaultCellStyle === b.defaultCellStyle;
}

function columnSignature(sheet: Worksheet, column: number, registry: StyleRegistry): ColumnSignature {
  const model = sheet.findColumn(column);
  if (!model) return { style: null, hidden: false, defaultCellStyle: null };
  const style = model.width !== Defaults.columnWidth ? registry.columnStyle(model.width) : null;
  const defaultCellStyle = model.styleId !== StylePool.defaultId ? registry.cellStyle(model.styleId) : null;
  return { style, hidden: model.hidden, defaultCellStyle };
}

/**
 * Write rows `0 … lastRow`. Rows holding a cell or a non-default row model are built one by one;
 * the stretches between them hold only what validations and merges put there, which is the same
 * for every row until one of those starts or ends, so each stretch is one repeated row. Then any
 * row identical to the one before it (no values, same cells, styles, validations and layout) is
 * folded into it, so a rule over B1:B1048576 is a handful of elements, not a million.
 */
function writeRows(table: XmlElement, w: SheetWriter, lastRow: number): void {
  const { sheet } = w;
  const cellColumns = new Map<number, number[]>();
  for (const [row, column] of sheet.enumerateCells()) {
    let list = cellColumns.get(row);
    if (!list) cellColumns.set(row, (list = []));
    list.push(column);
  }
  for (const list of cellColumns.values()) list.sort((a, b) => a - b);

  const special = new Set<number>(cellColumns.keys());
  for (const [row, model] of sheet.rows) if (!model.isDefault) special.add(row);
  const specialRows = [...special].sort((a, b) => a - b);

  const boundarySet = new Set<number>();
  for (const v of w.validations) boundarySet.add(v.range.top).add(v.range.bottom + 1);
  for (const m of sheet.merges)
    boundarySet
      .add(m.range.top)
      .add(m.range.top + 1)
      .add(m.range.bottom + 1);
  const boundaries = [...boundarySet].sort((a, b) => a - b);

  const rows = new RowAppender(table);
  let nextSpecial = 0;
  let nextBoundary = 0;
  let row = 0;
  // The layout of the last row written when it held no cells: a following empty row with the same
  // layout and no boundary between them is just one more repeat, with nothing to build.
  let lastEmptyLayout: string | null = null;
  while (row <= lastRow) {
    while (nextSpecial < specialRows.length && specialRows[nextSpecial]! < row) nextSpecial++;
    while (nextBoundary < boundaries.length && boundaries[nextBoundary]! <= row) nextBoundary++;
    const atBoundary = nextBoundary > 0 && boundaries[nextBoundary - 1] === row;
    if (specialRows[nextSpecial] === row) {
      const columns = cellColumns.get(row);
      const layout = columns ? null : rowLayout(sheet.findRow(row));
      if (layout !== null && layout === lastEmptyLayout && !atBoundary && rows.extend(1)) {
        row++;
        continue;
      }
      rows.append(writeRow(w, row, columns ?? []), 1);
      lastEmptyLayout = layout;
      row++;
      continue;
    }
    lastEmptyLayout = null;
    const until = Math.min(
      specialRows[nextSpecial] ?? Infinity,
      boundaries[nextBoundary] ?? Infinity,
      lastRow + 1,
    );
    rows.append(writeRow(w, row, []), until - row);
    row = until;
  }
}

/** Appends rows to a table, folding each into the previous one when they are interchangeable. */
class RowAppender {
  private readonly table: XmlElement;
  private last: XmlElement | null = null;
  private lastSignature: string | null = null;
  private lastRepeat = 0;

  constructor(table: XmlElement) {
    this.table = table;
  }

  /** Repeat the last row `count` more times; false when there is no row to repeat. */
  extend(count: number): boolean {
    if (!this.last) return false;
    this.lastRepeat += count;
    this.last.setAttr(Table('number-rows-repeated'), this.lastRepeat);
    return true;
  }

  append(row: XmlElement, repeat: number): void {
    const signature = rowSignature(row);
    if (this.last && signature !== null && signature === this.lastSignature) {
      this.lastRepeat += repeat;
      this.last.setAttr(Table('number-rows-repeated'), this.lastRepeat);
      return;
    }
    if (repeat > 1) row.setAttr(Table('number-rows-repeated'), repeat);
    this.table.add(row);
    this.last = row;
    this.lastSignature = signature;
    this.lastRepeat = repeat;
  }
}

/** What decides how an empty row is written: its height, visibility and default style. */
function rowLayout(
  model:
    | {
        readonly height: number;
        readonly manuallyHidden: boolean;
        readonly filtered: boolean;
        readonly styleId: number;
      }
    | undefined,
): string {
  return model ? `${model.height}|${model.manuallyHidden}|${model.filtered}|${model.styleId}` : '';
}

/**
 * A row's identity for folding: null when any cell holds content (values are never folded) or
 * anchors a merge (each anchor starts its own merge; repeated, LibreOffice keeps only the first).
 */
function rowSignature(row: XmlElement): string | null {
  let signature = attributeSignature(row, 'number-rows-repeated');
  for (const cell of row.elements()) {
    if (cell.nodes.length > 0 || isSpanned(cell)) return null;
    signature += `|${cell.local}${attributeSignature(cell, '')}`;
  }
  return signature;
}

function isSpanned(cell: XmlElement): boolean {
  return (
    cell.attr(Table('number-columns-spanned')) !== null || cell.attr(Table('number-rows-spanned')) !== null
  );
}

function attributeSignature(element: XmlElement, skip: string): string {
  let s = '';
  for (const a of element.attributes) {
    if (a.name.local === skip && a.name.uri === Table.uri) continue;
    s += ` ${a.name.uri}:${a.name.local}=${a.value}`;
  }
  return s;
}

/**
 * One row: its layout, then its cells left to right. Columns are visited in segments — a stored
 * cell, or a run with the same validation and merge state — so a blank row costs a few elements
 * however wide the sheet. Neighbouring blank cells that came out identical are folded into one
 * repeated cell.
 */
function writeRow(w: SheetWriter, row: number, cellColumns: readonly number[]): XmlElement {
  const { sheet, registry } = w;
  const element = new XmlElement(Table('table-row'));
  const rowModel = sheet.findRow(row);
  if (rowModel && rowModel.height !== Defaults.rowHeight) {
    element.setAttr(Table('style-name'), registry.rowStyle(rowModel.height));
  }
  // A filter's rows are `filter`, so LibreOffice shows them again when the filter is cleared; a row
  // both filtered and hidden by hand is written as filtered, as LibreOffice does.
  if (rowModel?.filtered) element.setAttr(Table('visibility'), 'filter');
  else if (rowModel?.manuallyHidden) element.setAttr(Table('visibility'), 'collapse');
  // A whole formatted row travels the same way a formatted column does — one attribute.
  if (rowModel && rowModel.styleId !== StylePool.defaultId) {
    element.setAttr(Table('default-cell-style-name'), registry.cellStyle(rowModel.styleId));
  }

  const validations = w.validations.filter((v) => v.range.top <= row && row <= v.range.bottom);
  const merges = w.merges.at(row);
  const validationAt = (column: number): string | null => {
    for (let i = validations.length - 1; i >= 0; i--) {
      const v = validations[i]!;
      if (v.range.left <= column && column <= v.range.right) return v.name;
    }
    return null;
  };
  const mergeAt = (column: number): MergedRegion | null =>
    merges.find((m) => m.range.left <= column && column <= m.range.right) ?? null;

  // Every column where what is written may change.
  const breaks = new Set<number>([0, w.lastColumn + 1]);
  for (const c of cellColumns) breaks.add(c).add(c + 1);
  for (const v of validations) breaks.add(v.range.left).add(v.range.right + 1);
  for (const m of merges)
    breaks
      .add(m.range.left)
      .add(m.range.left + 1)
      .add(m.range.right + 1);
  const points = [...breaks].filter((c) => c <= w.lastColumn + 1).sort((a, b) => a - b);

  const cells: XmlElement[] = [];
  for (let i = 0; i + 1 < points.length; i++) {
    const start = points[i]!;
    const width = points[i + 1]! - start;
    const merge = mergeAt(start);
    const covered = merge !== null && merge.isCovered(row, start);
    const validationName = validationAt(start);
    const cell = sheet.find(row, start);
    if (cell) {
      cells.push(writeCell(w, cell, covered ? null : merge, row, start, validationName, covered));
      continue;
    }
    if (merge && !covered) {
      // The anchor of a merge with no content still carries the spans.
      cells.push(writeCell(w, null, merge, row, start, validationName, false));
      continue;
    }
    const blank = new XmlElement(Table(covered ? 'covered-table-cell' : 'table-cell'));
    if (validationName !== null) blank.setAttr(Table(N.contentValidationName), validationName);
    if (width > 1) blank.setAttr(Table('number-columns-repeated'), width);
    cells.push(blank);
  }

  // Fold neighbouring empty cells that are alike (same kind, style and validation).
  let previous: XmlElement | null = null;
  let previousSignature: string | null = null;
  for (const cell of cells) {
    const signature =
      cell.nodes.length === 0 && !isSpanned(cell)
        ? cell.local + attributeSignature(cell, 'number-columns-repeated')
        : null;
    if (previous && signature !== null && signature === previousSignature) {
      const total = repeatOf(previous) + repeatOf(cell);
      previous.setAttr(Table('number-columns-repeated'), total);
      continue;
    }
    element.add(cell);
    previous = cell;
    previousSignature = signature;
  }
  return element;
}

function repeatOf(cell: XmlElement): number {
  return Number(cell.attr(Table('number-columns-repeated')) ?? '1');
}

function writeCell(
  w: SheetWriter,
  cell: Cell | null,
  merge: MergedRegion | null,
  row: number,
  column: number,
  validationName: string | null,
  covered: boolean,
): XmlElement {
  // A covered cell keeps its content too: LibreOffice shows it again when the merge is undone.
  const element = new XmlElement(Table(covered ? 'covered-table-cell' : 'table-cell'));

  if (cell && cell.styleId !== StylePool.defaultId) {
    element.setAttr(Table('style-name'), w.registry.cellStyle(cell.styleId));
  }

  if (validationName !== null) element.setAttr(Table(N.contentValidationName), validationName);

  if (merge && merge.range.top === row && merge.range.left === column) {
    if (merge.range.columnCount > 1)
      element.setAttr(Table('number-columns-spanned'), merge.range.columnCount);
    if (merge.range.rowCount > 1) element.setAttr(Table('number-rows-spanned'), merge.range.rowCount);
  }

  if (cell) {
    const preserved = cell.unmodeled;
    for (const child of preserved?.before ?? []) element.add(child.clone());
    const checkbox = covered ? undefined : w.checkboxes.get(cellKey(row, column));
    if (checkbox !== undefined) element.add(checkboxControl(w.sheet, row, column, checkbox));
    writeCellContent(element, cell, w.registry);
    for (const child of preserved?.after ?? []) element.add(child.clone());
  }
  return element;
}

function writeCellContent(element: XmlElement, cell: Cell, registry: StyleRegistry): void {
  if (cell.formula !== null) {
    element.setAttr(Table('formula'), OpenFormulaTranslator.toOpenFormula(cell.formula, cell.formulaSource));
  }

  const value = cell.value;
  const format = registry.styles.get(cell.styleId).numberFormat;
  const kind = NumberFormats.kind(format);
  const isTime = kind === 'time';
  const isDate = kind === 'date' || kind === 'dateTime';

  let lines: string[] = [];
  switch (value.kind) {
    case 'number':
      if (isTime) {
        // A time format stores the serial as an ISO-8601 duration so the whole value (days *and*
        // time-of-day fraction) survives, including elapsed times past 24h.
        element.setAttr(Office('value-type'), 'time');
        element.setAttr(Office('time-value'), durationText(value.value));
      } else if (isDate) {
        element.setAttr(Office('value-type'), 'date');
        element.setAttr(Office('date-value'), dateValueText(value.value));
      } else if (kind === 'currency') {
        element.setAttr(Office('value-type'), 'currency');
        element.setAttr(Office('currency'), currencyCode(format.code));
        element.setAttr(Office('value'), floatValueText(value.value));
      } else if (kind === 'percent') {
        element.setAttr(Office('value-type'), 'percentage');
        element.setAttr(Office('value'), floatValueText(value.value));
      } else {
        element.setAttr(Office('value-type'), 'float');
        element.setAttr(Office('value'), floatValueText(value.value));
      }
      lines = [NumberFormatter.format(value, format.code, INVARIANT_LOCALE)];
      break;
    case 'boolean':
      element.setAttr(Office('value-type'), 'boolean');
      element.setAttr(Office('boolean-value'), value.value ? 'true' : 'false');
      lines = [value.value ? 'TRUE' : 'FALSE'];
      break;
    case 'text':
      element.setAttr(Office('value-type'), 'string');
      // Text a file marked up is written back as it was until it is edited.
      if (
        cell.richText !== null &&
        cell.richText.text === value.value &&
        cell.richText.hyperlink === (cell.hyperlink?.target ?? null)
      ) {
        for (const paragraph of cell.richText.paragraphs) element.add(paragraph.clone());
        return;
      }
      // One text:p per line: ODF has no newline inside a paragraph (a conformant reader collapses
      // it to a space). The reader joins paragraphs with "\n".
      // A pasted Windows line end ("\r\n") or a lone "\r" is a line break too, not a stray character.
      lines = value.value.replace(/\r\n?/g, '\n').split('\n');
      break;
    case 'error':
      // As LibreOffice writes it: an empty string value marked as an error, the token as display.
      element.setAttr(Office('value-type'), 'string');
      element.setAttr(Office('string-value'), '');
      element.setAttr(Calcext('value-type'), 'error');
      lines = [value.error.token];
      break;
    case 'empty':
      break;
  }

  lines.forEach((line, i) => {
    const paragraph = new XmlElement(Text('p'));
    // The hyperlink wraps the first paragraph's text, as it always has.
    const target =
      cell.hyperlink && i === 0
        ? new XmlElement(Text('a')).setAttr(Xlink('href'), cell.hyperlink.target)
        : null;
    appendParagraphText(target ?? paragraph, line);
    if (target) {
      // An empty text node marks the paragraph as mixed content so the writer does not indent it;
      // indentation inside text:p is real cell text to a reader.
      paragraph.add(new XmlText(''), target);
    }
    element.add(paragraph);
  });
}

const CURRENCY_BY_SYMBOL: Readonly<Record<string, string>> = { $: 'USD', '€': 'EUR', '£': 'GBP', '¥': 'JPY' };

/**
 * The ISO 4217 code for `office:currency`, read from the format's symbol: `[$€-407]`, `[$USD]`,
 * `"$"` or a bare `€`. Null (the attribute is optional) when the code names no currency we know.
 */
function currencyCode(code: string): string | null {
  const tagged = /\[\$([^\]-]+)/.exec(code)?.[1];
  const symbol = tagged ?? /[$€£¥]/.exec(code)?.[0];
  if (symbol === undefined) return null;
  if (/^[A-Z]{3}$/.test(symbol)) return symbol;
  return CURRENCY_BY_SYMBOL[symbol] ?? null;
}

/**
 * An ODF `office:value` (`xsd:double`). Non-finite values spell `INF`/`-INF`/`NaN`; "Infinity" is
 * rejected by conformant readers.
 */
function floatValueText(n: number): string {
  if (Number.isNaN(n)) return 'NaN';
  if (n === Infinity) return 'INF';
  if (n === -Infinity) return '-INF';
  return roundTripNumber(n);
}

/**
 * An ODF `office:date-value`. Pure dates stay `yyyy-MM-dd`; values with a time-of-day fraction emit
 * the full ISO-8601 timestamp. Out-of-range serials fall back to a plain date string.
 */
function dateValueText(serial: number): string {
  const when = DateConversion.tryToDateTime(serial, false);
  if (!when) return datePart(DateConversion.dateFromSerial(serial));
  const midnight = when.hour === 0 && when.minute === 0 && when.second === 0 && when.millisecond === 0;
  if (midnight) return datePart(when);
  // Sub-second time is part of the value: 12:00:00.5 must not come back as 12:00:00.
  const fraction = when.millisecond === 0 ? '' : '.' + pad(when.millisecond, 3).replace(/0+$/, '');
  return `${datePart(when)}T${pad(when.hour, 2)}:${pad(when.minute, 2)}:${pad(when.second, 2)}${fraction}`;
}

function datePart(d: CivilDateTime): string {
  return `${pad(d.year, 4)}-${pad(d.month, 2)}-${pad(d.day, 2)}`;
}

function pad(n: number, width: number): string {
  return String(n).padStart(width, '0');
}

const TICKS_PER_DAY = 864_000_000_000;
const TICKS_PER_SECOND = 10_000_000;

/**
 * `XmlConvert.ToString(TimeSpan.FromDays(days))`: an `xsd:duration` such as `P1DT6H`, `PT12H`,
 * `PT0.5S`, or `PT0S`, at tick (100 ns) precision with a trimmed seven-digit fraction.
 */
export function durationText(days: number): string {
  if (!Number.isFinite(days)) return 'PT0S';
  let ticks = Math.round(days * TICKS_PER_DAY);
  let sign = '';
  if (ticks < 0) {
    sign = '-';
    ticks = -ticks;
  }

  const d = Math.floor(ticks / TICKS_PER_DAY);
  let rest = ticks - d * TICKS_PER_DAY;
  const h = Math.floor(rest / (3600 * TICKS_PER_SECOND));
  rest -= h * 3600 * TICKS_PER_SECOND;
  const m = Math.floor(rest / (60 * TICKS_PER_SECOND));
  rest -= m * 60 * TICKS_PER_SECOND;
  const s = Math.floor(rest / TICKS_PER_SECOND);
  const fraction = rest - s * TICKS_PER_SECOND;

  let text = sign + 'P';
  if (d !== 0) text += `${d}D`;
  if (h !== 0 || m !== 0 || s !== 0 || fraction !== 0) {
    text += 'T';
    if (h !== 0) text += `${h}H`;
    if (m !== 0) text += `${m}M`;
    if (s !== 0 || fraction !== 0) {
      text += String(s);
      if (fraction !== 0) text += '.' + String(fraction).padStart(7, '0').replace(/0+$/, '');
      text += 'S';
    }
  } else if (d === 0) {
    text += 'T0S';
  }
  return text;
}

/**
 * The workbook's global names: every modeled named range, then each preserved entry (a named
 * expression such as `Rate = of:=0.07`) whose name the model does not now hold itself.
 */
function addNamedExpressions(
  spreadsheet: XmlElement,
  workbook: Workbook,
  preserved: readonly XmlElement[],
): void {
  const global = workbook.namedRanges.filter((n) => n.scope === null || n.scope === undefined);
  const fallbackSheet =
    workbook.worksheets.length > 0 ? workbook.worksheets[0]!.name : Defaults.firstSheetName;
  const container = namedExpressionsElement(global, preserved, fallbackSheet);
  if (container) spreadsheet.add(container);
}

/** A sheet's own names (`scope` = the sheet), written inside its table as ODF scopes them. */
function addLocalNamedExpressions(
  table: XmlElement,
  workbook: Workbook,
  sheet: Worksheet,
  preserved: readonly XmlElement[],
): void {
  const key = sheet.name.toLowerCase();
  const local = workbook.namedRanges.filter((n) => n.scope?.toLowerCase() === key);
  const container = namedExpressionsElement(local, preserved, sheet.name);
  if (container) table.add(container);
}

function namedExpressionsElement(
  modeled: readonly NamedRange[],
  preserved: readonly XmlElement[],
  fallbackSheet: string,
): XmlElement | null {
  const container = new XmlElement(Table('named-expressions'));
  const taken = new Set<string>();
  for (const named of modeled) {
    taken.add(named.name.toLowerCase());
    container.add(
      new XmlElement(Table('named-range'))
        .setAttr(Table('name'), named.name)
        .setAttr(Table('cell-range-address'), odfRangeAddress(named.target, fallbackSheet))
        .setAttr(Table('base-cell-address'), odfCellAddress(named.target.topLeft, fallbackSheet)),
    );
  }
  for (const element of preserved) {
    const name = element.attr(Table('name'));
    if (name !== null && taken.has(name.toLowerCase())) continue;
    container.add(element.clone());
  }
  return container.hasElements ? container : null;
}
