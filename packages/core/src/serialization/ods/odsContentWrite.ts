import { Defaults } from '../../defaults.ts';
import { OpenFormulaTranslator } from '../../formulas/openFormulaTranslator.ts';
import type { Cell } from '../../model/cell.ts';
import type { MergedRegion } from '../../model/mergedRegion.ts';
import type { CellRange } from '../../model/cellRange.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { cellKey } from '../../model/sheetDimensions.ts';
import { StylePool } from '../../styling/stylePool.ts';
import { parseXmlElement, XmlDocument, XmlElement, XmlText } from '../../xml/index.ts';
import { OdsCharts } from './odsCharts.ts';
import { checkboxCells, checkboxControl, checkboxForm, Form, formsElement, Ooo } from './odsCheckboxes.ts';
import { appendParagraphText } from './odsContentShared.ts';
import { writeCellValue } from './odsCellValue.ts';
import { addDatabaseRanges } from './odsDatabaseRanges.ts';
import { addLocalNamedExpressions, addNamedExpressions } from './odsNamedExpressions.ts';
import type { ForeignXml } from './odsPackage.ts';
import { LineContainers } from './odsLineContainers.ts';
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
    // Schema order: office:scripts, office:font-face-decls, office:automatic-styles, office:body.
    .add(
      preserved?.scripts?.clone(),
      preserved?.fontFaces?.clone(),
      autoStyles,
      new XmlElement(Office('body'), spreadsheet),
    );

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

  // Groups and print titles are written in full, even over lines that hold nothing.
  const { outline } = sheet;
  const columnContainers = new LineContainers(table, 'column', outline.columnGroups, outline.headerColumns);
  const rowContainers = new LineContainers(table, 'row', outline.rowGroups, outline.headerRows);
  lastColumn = Math.max(lastColumn, columnContainers.lastLine);
  lastRow = Math.max(lastRow, rowContainers.lastLine);

  writeColumns(columnContainers, sheet, lastColumn, registry);
  const merges = new MergeSweep(sheet.merges);
  writeRows(rowContainers, { sheet, registry, validations, lastColumn, checkboxes, merges }, lastRow);

  // A chart sheet embeds its chart as an ODF draw frame (the object was written by writeObjects).
  if (sheet.isChartSheet) OdsCharts.injectFrame(table, sheet);

  addLocalNamedExpressions(table, workbook, sheet, preserved?.namedExpressions ?? []);
  for (const child of preserved?.after ?? []) table.add(child.clone());
  return table;
}

function writeColumns(
  containers: LineContainers,
  sheet: Worksheet,
  lastColumn: number,
  registry: StyleRegistry,
): void {
  // One element per column up to the used width, coalescing identical neighbours inside one group.
  const boundaries = new Set(containers.boundaries);
  let column = 0;
  while (column <= lastColumn) {
    const signature = columnSignature(sheet, column, registry);
    let repeat = 1;
    while (
      column + repeat <= lastColumn &&
      !boundaries.has(column + repeat) &&
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

    containers.at(column).add(element);
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
  const style =
    model.width !== Defaults.columnWidth || model.pageBreak
      ? registry.columnStyle(model.width, model.pageBreak)
      : null;
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
function writeRows(containers: LineContainers, w: SheetWriter, lastRow: number): void {
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
  for (const b of containers.boundaries) boundarySet.add(b);
  const boundaries = [...boundarySet].sort((a, b) => a - b);

  const rows = new RowAppender(containers);
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
      rows.append(writeRow(w, row, columns ?? []), row, 1);
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
    rows.append(writeRow(w, row, []), row, until - row);
    row = until;
  }
}

/**
 * Appends rows to a table (or the group they are in), folding each into the previous one when they
 * are interchangeable and in the same group.
 */
class RowAppender {
  private readonly containers: LineContainers;
  private container: XmlElement | null = null;
  private last: XmlElement | null = null;
  private lastSignature: string | null = null;
  private lastRepeat = 0;

  constructor(containers: LineContainers) {
    this.containers = containers;
  }

  /** Repeat the last row `count` more times; false when there is no row to repeat. */
  extend(count: number): boolean {
    if (!this.last) return false;
    this.lastRepeat += count;
    this.last.setAttr(Table('number-rows-repeated'), this.lastRepeat);
    return true;
  }

  append(row: XmlElement, position: number, repeat: number): void {
    const signature = rowSignature(row);
    const container = this.containers.at(position);
    const sameContainer = container === this.container;
    this.container = container;
    if (this.last && sameContainer && signature !== null && signature === this.lastSignature) {
      this.lastRepeat += repeat;
      this.last.setAttr(Table('number-rows-repeated'), this.lastRepeat);
      return;
    }
    if (repeat > 1) row.setAttr(Table('number-rows-repeated'), repeat);
    container.add(row);
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
        readonly pageBreak: boolean;
        readonly styleId: number;
      }
    | undefined,
): string {
  return model
    ? `${model.height}|${model.manuallyHidden}|${model.filtered}|${model.pageBreak}|${model.styleId}`
    : '';
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
  if (rowModel && (rowModel.height !== Defaults.rowHeight || rowModel.pageBreak)) {
    element.setAttr(Table('style-name'), registry.rowStyle(rowModel.height, rowModel.pageBreak));
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
    // A cell without a style of its own shows its row's or column's number format, so its value is
    // typed by that format too: a date under a date column is an office:date-value, as LibreOffice
    // writes it, not a float whose text is the serial.
    writeCellContent(element, cell, w.registry, w.sheet.effectiveStyleId(row, column));
    for (const child of preserved?.after ?? []) element.add(child.clone());
  }
  return element;
}

function writeCellContent(element: XmlElement, cell: Cell, registry: StyleRegistry, styleId: number): void {
  if (cell.formula !== null) {
    element.setAttr(Table('formula'), OpenFormulaTranslator.toOpenFormula(cell.formula, cell.formulaSource));
  }

  const value = cell.value;
  const format = registry.styles.get(styleId).numberFormat;

  let lines: string[] = [];
  switch (value.kind) {
    case 'number':
    case 'boolean':
    case 'error':
      lines = writeCellValue(element, value, format);
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
