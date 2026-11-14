import { DateConversion } from '../../formulas/evaluation/dateConversion.ts';
import { OpenFormulaTranslator } from '../../formulas/openFormulaTranslator.ts';
import type { CellRichText, UnmodeledCellContent, UnparsedFormulaSource } from '../../model/cell.ts';
import { CellAddress } from '../../model/cellAddress.ts';
import { CellRange } from '../../model/cellRange.ts';
import type { Hyperlink } from '../../model/hyperlink.ts';
import { MergedRegion } from '../../model/mergedRegion.ts';
import { SheetDimensions } from '../../model/sheetDimensions.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { StylePool } from '../../styling/stylePool.ts';
import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import {
  appendXmlText,
  openXmlElement,
  sameName,
  serializeXmlElement,
  XmlComment,
  XmlElement,
  XmlEventReader,
  XmlProcessingInstruction,
  type XmlAttribute,
  type XmlDocument,
  type XmlEventHandler,
  type XmlName,
} from '../../xml/index.ts';
import { parseInvariantFloat } from './mapping/odsCellStyles.ts';
import { ChartSheetProbe, OdsCharts } from './odsCharts.ts';
import {
  applyCheckboxes,
  CheckboxCollector,
  isTakenControl,
  withoutCheckboxes,
  type ReadCheckboxes,
} from './odsCheckboxes.ts';
import { readDatabaseRanges } from './odsDatabaseRanges.ts';
import type { ForeignXml, OdsPackage } from './odsPackage.ts';
import { OdsPreserved } from './odsPreserved.ts';
import { paragraphText } from './odsContentShared.ts';
import { OdsStyleIndex } from './odsStyleIndex.ts';
import {
  assembleValidations,
  readContentValidationRules,
  recordValidationRect,
  type ValidationCells,
} from './odsValidation.ts';
import { Calcext, Draw, OdsXmlNames as N, Office, Table, Text, Xlink } from './odsXmlNames.ts';

const MAX_COLUMNS = SheetDimensions.maxColumns;
const MAX_ROWS = SheetDimensions.maxRows;

/** A styled blank run larger than this is applied as a row or column style instead of per cell. */
const STYLED_RUN_CELL_LIMIT = 1 << 16;
/** The most cells one repeated run may create; past it the file is refused, not half-read. */
const CONTENT_RUN_CELL_LIMIT = 1 << 22;
/**
 * The most cells repeated runs may create in one workbook, all sheets together: as many as one run
 * may cover, since a few rows of runs at that cap would fill memory from a file of a few hundred
 * bytes. LibreOffice repeats only identical neighbours (and Cellestra reads big blank styled runs
 * as row or column formats), so a real file stays far below it.
 */
const REPEATED_CELL_BUDGET = CONTENT_RUN_CELL_LIMIT;
/**
 * The most merges repeated spanned cells may create in one workbook. LibreOffice never repeats a
 * cell that starts a merge, and neither does Cellestra, so only a corrupt or hostile file gets near.
 */
const REPEATED_MERGE_BUDGET = 1 << 15;

/** The cells and merges repeats have created so far, shared by every sheet of a workbook. */
interface RepeatBudget {
  cells: number;
  merges: number;
}

/** The width of a LibreOffice sheet before 7.4: a run that reaches it reaches the end of the row. */
const LEGACY_COLUMNS = 1024;

/** Thrown when a file asks for more cells than can be held (a hostile or corrupt repeat count). */
export class OdsContentTooLargeError extends Error {
  override readonly name = 'OdsContentTooLargeError';
}

/** The row and column containers ODF allows a grid to be nested in (ODF 9.1.2, 9.1.6). */
const ROW_GROUPS: readonly XmlName[] = [
  Table('table-header-rows'),
  Table('table-row-group'),
  Table('table-rows'),
];
const COLUMN_GROUPS: readonly XmlName[] = [
  Table('table-header-columns'),
  Table('table-column-group'),
  Table('table-columns'),
];

interface ReadContext {
  readonly styles: OdsStyleIndex;
  readonly validationCells: ValidationCells;
  /**
   * The sheet is a chart sheet, so its chart frame is modeled (not preserved) content. Settled
   * before the first cell that could hold the frame is read.
   */
  chartSheet: boolean;
  /** Checkbox controls the model takes over; their cell-anchored `draw:control`s are dropped. */
  readonly checkboxIds: ReadonlySet<string>;
  /** The table's merges, gathered in reading order and added once the grid is read. */
  readonly merges: MergedRegion[];
  /** What repeats have created in the whole workbook. */
  readonly repeats: RepeatBudget;
}

/**
 * Read `content.xml`, given as text in one or more chunks, into `workbook`, and return the
 * document's outline for the page setup: the root with its automatic styles, and each table
 * element with its attributes but none of its content (null when there is no root element).
 *
 * The document is read as a stream: each row element is built, read into the sheet and dropped,
 * so a large file never has its whole tree in memory. Unmodeled `office:spreadsheet` children are
 * appended to `foreign` as verbatim XML (with whether they came before the first table) so the
 * writer can re-emit them in place; unmodeled table, cell and named-expression content is kept
 * through `OdsPreserved`. `stylesDocument` is styles.xml, whose common and data styles cells may
 * name. Throws `XmlParseError` when the text is not well-formed.
 */
export function readContentXml(
  chunks: Iterable<string>,
  workbook: Workbook,
  foreign: ForeignXml[],
  pkg: OdsPackage | null = null,
  stylesDocument: XmlDocument | null = null,
): XmlElement {
  const content = new ContentReader(workbook, foreign, pkg, stylesDocument?.root ?? null);
  const reader = new XmlEventReader(content);
  for (const chunk of chunks) reader.write(chunk);
  reader.close();
  return content.finish();
}

/** The row and column groups' leaves and groups, which make up a table's grid. */
const isRowGroup = (name: XmlName): boolean => ROW_GROUPS.some((g) => sameName(g, name));
const isColumnGroup = (name: XmlName): boolean => COLUMN_GROUPS.some((g) => sameName(g, name));
const TABLE_ROW = Table('table-row');
const TABLE_COLUMN = Table('table-column');

/**
 * Where the reader is in the document. `element` frames build a tree: the outline (`root`,
 * `body`, `spreadsheet`) or a `subtree` that is handed to `done` when it closes. The other
 * frames build nothing: a table, a row or column group in it, or content that is not read.
 */
type Frame =
  | {
      readonly kind: 'element';
      readonly element: XmlElement;
      readonly role: 'root' | 'body' | 'spreadsheet' | 'subtree';
      readonly done: ((e: XmlElement) => void) | null;
    }
  | { readonly kind: 'table' | 'rows' | 'columns'; readonly table: TableReader }
  | { readonly kind: 'skip' };

const SKIP: Frame = { kind: 'skip' };

/**
 * Turns the events of `content.xml` into the workbook. The outline before `office:body`
 * (automatic styles, font faces) is built as a tree; in the first `office:spreadsheet`, each
 * child other than a table is built and read when it closes, and each table is read by a
 * `TableReader` one row element at a time.
 */
class ContentReader implements XmlEventHandler {
  private readonly workbook: Workbook;
  private readonly foreign: ForeignXml[];
  private readonly pkg: OdsPackage | null;
  private readonly stylesRoot: XmlElement | null;
  private readonly frames: Frame[] = [];
  private root: XmlElement | null = null;
  private sawBody = false;
  private spreadsheet: XmlElement | null = null;
  private styleIndex: OdsStyleIndex | null = null;
  private seenTable = false;
  private validations: XmlElement | null = null;
  private readonly validationCells: ValidationCells = new Map();
  private readonly namedExpressions: XmlElement[] = [];
  private readonly checkboxes: ReadCheckboxes = { ids: new Set(), cells: [] };
  private readonly repeats: RepeatBudget = { cells: 0, merges: 0 };

  constructor(
    workbook: Workbook,
    foreign: ForeignXml[],
    pkg: OdsPackage | null,
    stylesRoot: XmlElement | null,
  ) {
    this.workbook = workbook;
    this.foreign = foreign;
    this.pkg = pkg;
    this.stylesRoot = stylesRoot;
  }

  /** The styles cells name: content.xml's automatic styles come before its body, so are complete. */
  private get styles(): OdsStyleIndex {
    this.styleIndex ??= new OdsStyleIndex(this.root!, this.stylesRoot, this.workbook.styles);
    return this.styleIndex;
  }

  declaration(): void {}

  openElement(name: XmlName, attributes: XmlAttribute[]): void {
    const top = this.frames.at(-1);
    if (!top) {
      this.root = openXmlElement(null, name, attributes);
      this.frames.push({ kind: 'element', element: this.root, role: 'root', done: null });
      return;
    }
    switch (top.kind) {
      case 'skip':
        this.frames.push(SKIP);
        return;
      case 'rows':
        if (sameName(name, TABLE_ROW)) this.subtree(name, attributes, (e) => top.table.row(e));
        else this.frames.push(isRowGroup(name) ? top : SKIP);
        return;
      case 'columns':
        if (sameName(name, TABLE_COLUMN)) this.subtree(name, attributes, (e) => top.table.column(e));
        else this.frames.push(isColumnGroup(name) ? top : SKIP);
        return;
      case 'table':
        this.openInTable(top.table, name, attributes);
        return;
      case 'element':
        this.openInOutline(top, name, attributes);
    }
  }

  private openInOutline(top: Frame & { kind: 'element' }, name: XmlName, attributes: XmlAttribute[]): void {
    switch (top.role) {
      case 'root': {
        const element = openXmlElement(top.element, name, attributes);
        const body = !this.sawBody && sameName(name, Office('body'));
        if (body) this.sawBody = true;
        this.frames.push({ kind: 'element', element, role: body ? 'body' : 'subtree', done: null });
        return;
      }
      case 'body':
        // Only the first office:spreadsheet of the body is read.
        if (this.spreadsheet || !sameName(name, Office('spreadsheet'))) {
          this.frames.push(SKIP);
          return;
        }
        this.spreadsheet = openXmlElement(top.element, name, attributes);
        this.frames.push({ kind: 'element', element: this.spreadsheet, role: 'spreadsheet', done: null });
        return;
      case 'spreadsheet':
        if (sameName(name, Table('table'))) {
          this.seenTable = true;
          // The outline keeps the table's attributes, which the page setup reads.
          const outline = openXmlElement(top.element, name, attributes);
          this.frames.push({ kind: 'table', table: this.startTable(outline) });
        } else {
          this.subtree(name, attributes, (e) => this.spreadsheetChild(e));
        }
        return;
      case 'subtree': {
        const element = openXmlElement(top.element, name, attributes);
        this.frames.push({ kind: 'element', element, role: 'subtree', done: null });
      }
    }
  }

  private openInTable(table: TableReader, name: XmlName, attributes: XmlAttribute[]): void {
    if (sameName(name, TABLE_ROW)) {
      table.gridStarted();
      this.subtree(name, attributes, (e) => table.row(e));
    } else if (sameName(name, TABLE_COLUMN)) {
      table.gridStarted();
      this.subtree(name, attributes, (e) => table.column(e));
    } else if (isRowGroup(name)) {
      table.gridStarted();
      this.frames.push({ kind: 'rows', table });
    } else if (isColumnGroup(name)) {
      table.gridStarted();
      this.frames.push({ kind: 'columns', table });
    } else {
      const afterGrid = table.hasGrid;
      this.subtree(name, attributes, (e) => table.child(e, afterGrid));
    }
  }

  /** Build a detached element and everything in it, and hand it to `done` when it closes. */
  private subtree(name: XmlName, attributes: XmlAttribute[], done: (e: XmlElement) => void): void {
    const element = openXmlElement(null, name, attributes);
    this.frames.push({ kind: 'element', element, role: 'subtree', done });
  }

  closeElement(): void {
    const frame = this.frames.pop();
    if (frame?.kind === 'element') {
      frame.element.trimToSize();
      frame.done?.(frame.element);
    } else if (frame?.kind === 'table') {
      frame.table.finish();
    }
  }

  text(value: string, cdata: boolean): void {
    const top = this.frames.at(-1);
    if (top?.kind === 'element') appendXmlText(top.element, value, cdata, true);
  }

  comment(value: string): void {
    const top = this.frames.at(-1);
    if (top?.kind === 'element') top.element.nodes.push(new XmlComment(value));
  }

  processingInstruction(target: string, body: string): void {
    const top = this.frames.at(-1);
    if (top?.kind === 'element') top.element.nodes.push(new XmlProcessingInstruction(target, body));
  }

  private startTable(outline: XmlElement): TableReader {
    const workbook = this.workbook;
    const name = outline.attr(Table('name')) ?? workbook.addSheet().name;
    const sheet = workbook.findSheet(name) ?? workbook.addSheetForLoad(name);
    const context: ReadContext = {
      styles: this.styles,
      validationCells: this.validationCells,
      chartSheet: false,
      checkboxIds: this.checkboxes.ids,
      merges: [],
      repeats: this.repeats,
    };
    return new TableReader(workbook, sheet, outline, context, this.pkg, this.checkboxes);
  }

  private spreadsheetChild(child: XmlElement): void {
    if (child.is(Table('named-expressions'))) {
      this.namedExpressions.push(...readNamedExpressions(child, this.workbook, null));
    } else if (child.is(Table(N.databaseRanges))) {
      readDatabaseRanges(child, this.workbook);
    } else if (child.is(Table(N.contentValidations))) {
      // Modeled: the rules are read at the end, once every cell has named the ones it uses.
      this.validations ??= child;
    } else {
      this.foreign.push({ beforeTables: !this.seenTable, xml: serializeXmlElement(child) });
    }
  }

  /** The document is complete: settle what needed every table, and return the outline. */
  finish(): XmlElement {
    const root = this.root!;
    if (!this.spreadsheet) return root;
    // Rule definitions are read last; each rule's target is rebuilt from the cells that named it.
    const container = new XmlElement(Office('spreadsheet'), this.validations);
    assembleValidations(this.workbook, readContentValidationRules(container), this.validationCells);
    applyCheckboxes(this.workbook, this.checkboxes);
    OdsPreserved.setWorkbook(this.workbook, {
      namedExpressions: this.namedExpressions,
      automaticStyles: this.styles.preservedAutomaticStyles,
      fontFaces: root.element(Office('font-face-decls'))?.clone() ?? null,
    });
    return root;
  }
}

/**
 * Reads one `table:table` as its children arrive. Grid rows go to a `TableGridReader`, except
 * while the table may still be a chart sheet or a checkbox still waits for its linked cell: those
 * change how a cell is read, so rows are held until they are settled (a chart sheet has no data,
 * and a checkbox's control sits in its own cell, so little is ever held). Every other child of the
 * table is kept, and read when the table ends.
 */
class TableReader {
  private readonly workbook: Workbook;
  private readonly sheet: Worksheet;
  private readonly outline: XmlElement;
  private readonly context: ReadContext;
  private readonly pkg: OdsPackage | null;
  private readonly grid: TableGridReader;
  private readonly chart = new ChartSheetProbe();
  private readonly checkboxes: CheckboxCollector;
  private readonly held: XmlElement[] = [];
  private readonly children: Array<{ readonly element: XmlElement; readonly afterGrid: boolean }> = [];
  private sawForms = false;
  private sawRow = false;
  hasGrid = false;

  constructor(
    workbook: Workbook,
    sheet: Worksheet,
    outline: XmlElement,
    context: ReadContext,
    pkg: OdsPackage | null,
    checkboxes: ReadCheckboxes,
  ) {
    this.workbook = workbook;
    this.sheet = sheet;
    this.outline = outline;
    this.context = context;
    this.pkg = pkg;
    this.grid = new TableGridReader(sheet, context);
    this.checkboxes = new CheckboxCollector(sheet.name, checkboxes);
  }

  gridStarted(): void {
    this.hasGrid = true;
  }

  column(element: XmlElement): void {
    this.grid.column(element);
  }

  row(element: XmlElement): void {
    this.sawRow = true;
    this.chart.row(element);
    this.checkboxes.row(element);
    // A chart sheet needs a probe to its end (pkg is how a chart is found, so none means none).
    const waiting = (this.pkg !== null && !this.chart.settled) || this.checkboxes.pending;
    if (waiting) {
      this.held.push(element);
      return;
    }
    this.readHeld();
    this.grid.row(element);
  }

  /** A child of the table outside its grid, once complete. */
  child(element: XmlElement, afterGrid: boolean): void {
    this.chart.nonGridChild();
    // The table's first office:forms names its checkboxes; it comes before the grid.
    if (!this.sawForms && element.is(Office('forms'))) {
      this.sawForms = true;
      if (!this.sawRow) this.checkboxes.forms(element);
    }
    this.children.push({ element, afterGrid });
  }

  private readHeld(): void {
    for (const element of this.held) this.grid.row(element);
    this.held.length = 0;
  }

  finish(): void {
    // A Cellestra chart sheet is a table holding nothing but the chart's frame in A1; the chart is
    // modeled and its frame regenerated. A chart anywhere else is preserved like any shape.
    this.context.chartSheet = OdsCharts.readChartSheet(this.chart, this.pkg, this.sheet) !== null;
    this.checkboxes.finish();
    this.readHeld();
    this.grid.finish();

    // Everything else the table carries is kept as it was: its attributes (protection, print
    // ranges, its table style) and its unmodeled children (shapes, forms, conditional formats).
    const before: XmlElement[] = [];
    const after: XmlElement[] = [];
    const namedExpressions: XmlElement[] = [];
    const ids = this.context.checkboxIds;
    for (const { element, afterGrid } of this.children) {
      if (element.is(Table('named-expressions'))) {
        namedExpressions.push(...readNamedExpressions(element, this.workbook, this.sheet.name));
        continue;
      }
      // The checkboxes the model took over are regenerated on save; the rest of the forms stay.
      const kept = withoutCheckboxes(element, ids);
      if (kept) (afterGrid ? after : before).push(kept);
    }
    const attributes = this.outline.attributes.filter(
      (a) => !(a.name.uri === Table.uri && a.name.local === 'name'),
    );
    if (attributes.length > 0 || before.length > 0 || after.length > 0 || namedExpressions.length > 0) {
      OdsPreserved.setTable(this.sheet, {
        attributes: attributes.map((a) => ({ name: a.name, value: a.value })),
        before,
        after,
        namedExpressions,
      });
    }
  }
}

function isHidden(element: XmlElement): boolean {
  const visibility = element.attr(Table('visibility'));
  return visibility === 'collapse' || visibility === 'filter';
}

function defaultStyleIdOf(element: XmlElement, styles: OdsStyleIndex): number | null {
  const name = element.attr(Table('default-cell-style-name'));
  if (name === null) return null;
  const id = styles.cellStyleId(name);
  return id === StylePool.defaultId ? null : id;
}

/** A run of styled blank cells down to the end of the sheet, applied as a column format. */
interface ColumnFormat {
  readonly from: number;
  readonly to: number;
  readonly styleId: number;
}

/** A row element waiting for the rows after it, which decide whether it is trailing or padding. */
interface HeldRow {
  readonly element: XmlElement;
  readonly row: number;
  readonly repeat: number;
}

/**
 * Consecutive inert rows laid out alike, held as one range: if nothing but them follows, they are
 * the sheet's padding. `trailingFrom` is where the row element that reaches the sheet's last row starts, if one does.
 */
interface HeldInertRun {
  readonly element: XmlElement;
  readonly layout: string;
  readonly from: number;
  to: number;
  trailingFrom: number | null;
}

/**
 * Reads a table's grid one `table:table-column` and `table:table-row` element at a time, in
 * document order, so the caller never needs the whole grid at once. A row is read once the next
 * row arrives (or the table ends), because whether it is the last row, or part of the blank
 * padding at the end of the sheet, changes how it is read.
 */
class TableGridReader {
  private readonly sheet: Worksheet;
  private readonly context: ReadContext;
  private nextColumn = 0;
  private nextRow = 0;
  private held: HeldRow | HeldInertRun | null = null;
  /** Column formats wait for the end of the table so the columns' own settings come first. */
  private readonly columnFormats: ColumnFormat[] = [];

  constructor(sheet: Worksheet, context: ReadContext) {
    this.sheet = sheet;
    this.context = context;
  }

  column(element: XmlElement): void {
    const column = this.nextColumn;
    if (column >= MAX_COLUMNS) return;
    const repeat = Math.min(readPositive(element, Table('number-columns-repeated')), MAX_COLUMNS - column);
    this.nextColumn += repeat;
    const styleName = element.attr(Table('style-name'));
    const hidden = isHidden(element);
    const width = styleName === null ? null : this.context.styles.columnWidth(styleName);
    // The column's default cell style: "this whole column is formatted like so".
    const defaultStyleId = defaultStyleIdOf(element, this.context.styles);
    if (width === null && !hidden && defaultStyleId === null) return;
    for (let c = column; c < column + repeat; c++) {
      const model = this.sheet.getOrCreateColumn(c);
      if (width !== null) model.width = width;
      if (defaultStyleId !== null) model.styleId = defaultStyleId;
      model.hidden = hidden;
    }
  }

  row(element: XmlElement): void {
    // Rows past the last one are not read, but still count when deciding what the padding is.
    const row = Math.min(this.nextRow, MAX_ROWS);
    const repeat = Math.min(readPositive(element, Table('number-rows-repeated')), MAX_ROWS - row);
    this.nextRow = row + repeat;
    const inert = rowIsInert(element);
    const layout = inert ? rowLayout(element) : '';
    const held = this.held;
    if (held && 'layout' in held && inert && held.layout === layout) {
      if (row + repeat >= MAX_ROWS && held.trailingFrom === null) held.trailingFrom = row;
      held.to = row + repeat;
      return;
    }
    if (held) this.readHeld(held, false);
    const trailingFrom = row + repeat >= MAX_ROWS ? row : null;
    this.held = inert
      ? { element, layout, from: row, to: row + repeat, trailingFrom }
      : { element, row, repeat };
  }

  /** The table is complete: read the rows still held, then the column formats and merges. */
  finish(): void {
    if (this.held) this.readHeld(this.held, true);
    this.held = null;
    for (const { from, to, styleId } of this.columnFormats) {
      for (let c = from; c < to; c++) this.sheet.getOrCreateColumn(c).styleId = styleId;
    }
    addMerges(this.sheet, this.context.merges);
  }

  private readHeld(held: HeldRow | HeldInertRun, last: boolean): void {
    if ('layout' in held) {
      // A run that ends the table is the padding; LibreOffice pads sheets with runs like 1,048,000
      // repeated rows of 1,024 blank cells, and the padding only ever carries the default row
      // height, so a plain run there is skipped whole. So is the row that reaches the last row.
      const trailingFrom = held.trailingFrom ?? held.to;
      this.readInertRows(held.element, held.from, trailingFrom, last);
      this.readInertRows(held.element, trailingFrom, held.to, true);
      return;
    }
    const { element, row, repeat } = held;
    if (repeat === 0) return;
    this.applyRowLayout(element, row, row + repeat);
    readCells(
      element,
      this.sheet,
      row,
      repeat,
      last || row + repeat >= MAX_ROWS,
      this.context,
      this.columnFormats,
    );
  }

  /** Inert rows `from … to - 1`: a plain run carries nothing (and with `skipHeight`, nor does its height). */
  private readInertRows(element: XmlElement, from: number, to: number, skipHeight: boolean): void {
    if (from >= to) return;
    const styles = this.context.styles;
    const styleName = element.attr(Table('style-name'));
    const height = styleName === null ? null : styles.rowHeight(styleName);
    if (!isHidden(element) && defaultStyleIdOf(element, styles) === null && (height === null || skipHeight))
      return;
    this.applyRowLayout(element, from, to);
  }

  private applyRowLayout(element: XmlElement, from: number, to: number): void {
    const styles = this.context.styles;
    const styleName = element.attr(Table('style-name'));
    const hidden = isHidden(element);
    const height = styleName === null ? null : styles.rowHeight(styleName);
    const defaultStyleId = defaultStyleIdOf(element, styles);
    if (height === null && !hidden && defaultStyleId === null) return;
    for (let r = from; r < to; r++) {
      const model = this.sheet.getOrCreateRow(r);
      if (height !== null) model.height = height;
      if (defaultStyleId !== null) model.styleId = defaultStyleId;
      model.hidden = hidden;
    }
  }
}

/** What makes padding rows alike: their style, visibility and default cell style. */
function rowLayout(e: XmlElement): string {
  return `${e.attr(Table('style-name'))}|${e.attr(Table('visibility'))}|${e.attr(Table('default-cell-style-name'))}`;
}

/** A blank cell that only says how many columns it repeats over. */
function cellIsInert(element: XmlElement): boolean {
  if (element.nodes.some((n) => n.kind === 'element')) return false;
  return element.attributes.every(
    (a) => a.name.local === 'number-columns-repeated' && a.name.uri === element.uri,
  );
}

function isCellElement(element: XmlElement): boolean {
  return element.is(Table('table-cell')) || element.is(Table('covered-table-cell'));
}

/** Every cell of the row is inert (so reading it changes nothing). */
function rowIsInert(row: XmlElement): boolean {
  for (const element of row.elements()) if (isCellElement(element) && !cellIsInert(element)) return false;
  return true;
}

interface ParsedCell {
  readonly value: CellValue;
  readonly formula: string | null;
  readonly formulaSource: UnparsedFormulaSource | null;
  readonly styleId: number | null;
  readonly hyperlink: Hyperlink | null;
  readonly preserved: UnmodeledCellContent | null;
  readonly richText: CellRichText | null;
}

/**
 * Read one row element's cells into rows `row … row + rowCount - 1` (a repeated row is read once).
 * Validations are recorded as rectangles and big styled blank runs become row or column styles, so
 * a rule or a format over a million rows costs one entry, not a million.
 */
function readCells(
  rowElement: XmlElement,
  sheet: Worksheet,
  row: number,
  rowCount: number,
  trailingRows: boolean,
  context: ReadContext,
  columnFormats: ColumnFormat[],
): void {
  const cells = rowElement.elements().filter(isCellElement);
  let column = 0;
  // The style every cell so far in the row has: undefined before the first, null once they differ.
  let rowStyleBefore: number | null | undefined = undefined;
  const passStyle = (styleId: number): void => {
    if (rowStyleBefore === undefined) rowStyleBefore = styleId;
    else if (rowStyleBefore !== styleId) rowStyleBefore = null;
  };
  for (let index = 0; index < cells.length; index++) {
    if (column >= MAX_COLUMNS) break;
    const element = cells[index]!;
    const repeat = Math.min(readPositive(element, Table('number-columns-repeated')), MAX_COLUMNS - column);
    if (cellIsInert(element)) {
      passStyle(StylePool.defaultId);
      column += repeat;
      continue;
    }

    const validationName = element.attr(Table(N.contentValidationName));
    if (validationName !== null) {
      recordValidationRect(
        context.validationCells,
        sheet.name,
        validationName,
        row,
        row + rowCount - 1,
        column,
        column + repeat - 1,
      );
    }

    const parsed = parseCell(element, context);
    const hasContent =
      parsed.value.kind !== 'empty' ||
      parsed.formula !== null ||
      parsed.hyperlink !== null ||
      parsed.preserved !== null;
    const styled = parsed.styleId !== null && parsed.styleId !== StylePool.defaultId;
    const area = repeat * rowCount;

    if (hasContent || styled) {
      const lastInRow = index === cells.length - 1;
      const big = area > STYLED_RUN_CELL_LIMIT;
      // A row or column format would style every cell outside the run too, so a run only
      // becomes one when it covers the whole row or column. A run too big to hold as cells is a
      // format anyway, as it always was: a lossy read beats refusing the file.
      const tooBigForCells = area > CONTENT_RUN_CELL_LIMIT;
      const reachesRowEnd = column + repeat >= LEGACY_COLUMNS;
      const rowFormat =
        !hasContent &&
        lastInRow &&
        (big || (column + repeat >= MAX_COLUMNS && repeat > 1)) &&
        ((reachesRowEnd && (rowStyleBefore === undefined || rowStyleBefore === parsed.styleId)) ||
          tooBigForCells);
      const columnRun = !rowFormat && !hasContent && big && trailingRows;
      // Columns formatted below a header: the cells above that a column format would restyle.
      const pins = columnRun ? pinsAbove(sheet, row, column, repeat, parsed.styleId!, context) : null;
      if (rowFormat) {
        // A styled run to the end of the row (LibreOffice writes one for a formatted whole row):
        // that is a row format, not 16384 cells.
        for (let r = row; r < row + rowCount; r++) sheet.getOrCreateRow(r).styleId = parsed.styleId!;
      } else if (columnRun && (pins !== null || tooBigForCells)) {
        // A styled run down to the end of the sheet: that is a column format. The few cells above
        // it that look otherwise keep their style; when there are too many to keep, a run too big
        // to hold as cells is a format anyway.
        if (pins !== null) pinCells(sheet, pins, context);
        columnFormats.push({ from: column, to: column + repeat, styleId: parsed.styleId! });
      } else {
        if (tooBigForCells) {
          throw new OdsContentTooLargeError(
            `Sheet '${sheet.name}' repeats one cell over ${area} positions; the most a run may cover is ${CONTENT_RUN_CELL_LIMIT}.`,
          );
        }
        if (area > 1) {
          context.repeats.cells += area;
          if (context.repeats.cells > REPEATED_CELL_BUDGET) {
            throw new OdsContentTooLargeError(
              `Sheet '${sheet.name}' repeats cells past ${REPEATED_CELL_BUDGET} in the workbook, the most repeats may create.`,
            );
          }
        }
        for (let r = row; r < row + rowCount; r++) {
          for (let c = column; c < column + repeat; c++) storeCell(sheet, r, c, parsed);
        }
      }
    }

    if (element.is(Table('table-cell'))) {
      const columnsSpanned = Math.min(
        readPositive(element, Table('number-columns-spanned')),
        MAX_COLUMNS - column,
      );
      const declaredRowsSpanned = readPositive(element, Table('number-rows-spanned'));
      if (columnsSpanned > 1 || declaredRowsSpanned > 1) {
        // Each repeated copy starts a merge, but a copy inside the span of the one before it
        // overlaps it and would be dropped by addMerges: step over the spans instead, so only the
        // merges that can stand are made.
        const rowStep = Math.min(declaredRowsSpanned, rowCount);
        const columnStep = Math.min(columnsSpanned, repeat);
        if (rowCount > 1 || repeat > 1) {
          context.repeats.merges += Math.ceil(rowCount / rowStep) * Math.ceil(repeat / columnStep);
          if (context.repeats.merges > REPEATED_MERGE_BUDGET) {
            throw new OdsContentTooLargeError(
              `Sheet '${sheet.name}' repeats merged cells past ${REPEATED_MERGE_BUDGET} in the workbook, the most repeats may create.`,
            );
          }
        }
        for (let r = row; r < row + rowCount; r += rowStep) {
          const rowsSpanned = Math.min(declaredRowsSpanned, MAX_ROWS - r);
          if (columnsSpanned <= 1 && rowsSpanned <= 1) continue;
          for (let c = column; c < column + repeat; c += columnStep) {
            const right = Math.min(c + columnsSpanned - 1, MAX_COLUMNS - 1);
            const range = new CellRange(new CellAddress(c, r), new CellAddress(right, r + rowsSpanned - 1));
            context.merges.push(new MergedRegion(range));
          }
        }
      }
    }
    passStyle(parsed.styleId ?? StylePool.defaultId);
    column += repeat;
  }
}

/** A cell above a column format and the style it keeps. */
interface Pin {
  readonly row: number;
  readonly column: number;
  readonly styleId: number;
}

/**
 * The cells of `column … column + width - 1` above `row` that a column format of `styleId` would
 * restyle, each with the style it has now; null when there are more than `STYLED_RUN_CELL_LIMIT`
 * (then the run is cheaper to hold as cells) or more than `CONTENT_RUN_CELL_LIMIT` to look at. A cell or row with a style of its own keeps it
 * anyway, so in the usual layout (a styled header row over formatted columns) there is nothing to
 * pin; a header in the default style is pinned to a plain style of its own.
 */
function pinsAbove(
  sheet: Worksheet,
  row: number,
  column: number,
  width: number,
  styleId: number,
  context: ReadContext,
): Pin[] | null {
  if (row * width > CONTENT_RUN_CELL_LIMIT) return null;
  const pins: Pin[] = [];
  for (let r = 0; r < row; r++) {
    if ((sheet.findRow(r)?.styleId ?? StylePool.defaultId) !== StylePool.defaultId) continue;
    for (let c = column; c < column + width; c++) {
      if ((sheet.find(r, c)?.styleId ?? StylePool.defaultId) !== StylePool.defaultId) continue;
      const now = sheet.findColumn(c)?.styleId ?? StylePool.defaultId;
      if (now === styleId) continue;
      if (pins.length === STYLED_RUN_CELL_LIMIT) return null;
      pins.push({
        row: r,
        column: c,
        styleId: now === StylePool.defaultId ? context.styles.plainCellStyleId() : now,
      });
    }
  }
  return pins;
}

/** Give each pinned cell its style; the new cells count against the workbook's repeat budget. */
function pinCells(sheet: Worksheet, pins: readonly Pin[], context: ReadContext): void {
  context.repeats.cells += pins.length;
  if (context.repeats.cells > REPEATED_CELL_BUDGET) {
    throw new OdsContentTooLargeError(
      `Sheet '${sheet.name}' repeats cells past ${REPEATED_CELL_BUDGET} in the workbook, the most repeats may create.`,
    );
  }
  for (const pin of pins) sheet.getOrCreate(pin.row, pin.column).styleId = pin.styleId;
}

/**
 * Add the merges read from a table, skipping any that overlaps one already added: a corrupt file
 * with two overlapping spans loses the later one, not the whole load. `merges` arrive in reading
 * order (by top row), so only the merges still open at a candidate's top can overlap it.
 */
function addMerges(sheet: Worksheet, merges: readonly MergedRegion[]): void {
  if (merges.length === 0) return;
  const accepted = [...sheet.merges];
  let open = [...accepted];
  for (const merge of merges) {
    const top = merge.range.top;
    if (open.some((m) => m.range.bottom < top)) open = open.filter((m) => m.range.bottom >= top);
    if (open.some((m) => m.range.intersects(merge.range))) continue;
    accepted.push(merge);
    open.push(merge);
  }
  sheet.replaceMerges(accepted);
}

function storeCell(sheet: Worksheet, row: number, column: number, parsed: ParsedCell): void {
  const cell = sheet.getOrCreate(row, column);
  cell.value = parsed.value;
  cell.formula = parsed.formula;
  cell.formulaSource = parsed.formulaSource;
  cell.hyperlink = parsed.hyperlink;
  if (parsed.styleId !== null) cell.styleId = parsed.styleId;
  if (parsed.preserved) cell.unmodeled = parsed.preserved;
  if (parsed.richText) cell.richText = parsed.richText;
}

function parseCell(element: XmlElement, context: ReadContext): ParsedCell {
  const formula = element.attr(Table('formula'));
  const styleName = element.attr(Table('style-name'));

  // Children other than the text paragraphs — a comment (office:annotation), a shape anchored to
  // the cell, a detective mark — are kept verbatim, on the side of the text they were on.
  const before: XmlElement[] = [];
  const after: XmlElement[] = [];
  const paragraphs: XmlElement[] = [];
  for (const child of element.elements()) {
    if (child.is(Text('p'))) {
      paragraphs.push(child);
      continue;
    }
    if (context.chartSheet && child.is(Draw('frame'))) continue;
    if (isTakenControl(child, context.checkboxIds)) continue;
    (paragraphs.length > 0 ? after : before).push(child.clone());
  }

  const value = readValue(element, paragraphs);
  const hyperlink = readHyperlink(paragraphs);
  const read = formula !== null ? OpenFormulaTranslator.read(formula) : null;
  return {
    value,
    formula: read?.formula ?? null,
    formulaSource: read?.source ?? null,
    styleId: styleName === null ? null : context.styles.cellStyleId(styleName),
    hyperlink,
    preserved: before.length > 0 || after.length > 0 ? { before, after } : null,
    richText: value.kind === 'text' ? readRichText(paragraphs, hyperlink) : null,
  };
}

/**
 * The paragraphs of a text cell when they carry markup the model flattens (spans, links, line
 * breaks, paragraph styles); null for plain paragraphs, which the writer reproduces as they were.
 */
function readRichText(paragraphs: readonly XmlElement[], hyperlink: Hyperlink | null): CellRichText | null {
  if (!paragraphs.some((p) => p.hasAttributes || p.hasElements)) return null;
  return {
    text: paragraphs.map(paragraphText).join('\n'),
    hyperlink: hyperlink?.target ?? null,
    paragraphs: paragraphs.map((p) => p.clone()),
  };
}

function readValue(element: XmlElement, paragraphs: readonly XmlElement[]): CellValue {
  if (element.attr(Calcext('value-type')) === 'error') {
    // LibreOffice writes an error result as an empty string marked as an error, the token as display.
    // A token we have no code for (Calc's "Err:502") stays the text it shows.
    const shown = paragraphs.map(paragraphText).join('\n').trim();
    const code = CellError.codeFromToken(shown);
    return code === null
      ? CellValue.text(shown)
      : CellValue.error(new CellError(code, `The file saved this result as ${shown}.`));
  }
  switch (element.attr(Office('value-type'))) {
    case 'float':
    case 'percentage':
    case 'currency': {
      const n = parseDouble(element.attr(Office('value')));
      return n === null ? CellValue.empty : CellValue.number(n);
    }
    case 'date': {
      const serial = parseDateValue(element.attr(Office('date-value')));
      return serial === null ? CellValue.empty : CellValue.number(serial);
    }
    case 'time': {
      const days = parseDuration(element.attr(Office('time-value')));
      return days === null ? CellValue.empty : CellValue.number(days);
    }
    case 'boolean': {
      // xsd:boolean is "true"/"false" or "1"/"0".
      const b = element.attr(Office('boolean-value'))?.trim();
      return CellValue.boolean(b === 'true' || b === '1');
    }
    case 'string': {
      // office:string-value, when present, is the value; the paragraphs are only its display.
      const value = element.attr(Office('string-value'));
      return CellValue.text(value ?? paragraphs.map(paragraphText).join('\n'));
    }
    default: {
      // No value type, but text: a string cell from a writer that left the type out (ODF 19.385
      // makes it optional), which LibreOffice reads as text too.
      const text = paragraphs.map(paragraphText).join('\n');
      return text.length > 0 ? CellValue.text(text) : CellValue.empty;
    }
  }
}

function readHyperlink(paragraphs: readonly XmlElement[]): Hyperlink | null {
  for (const paragraph of paragraphs) {
    const href = paragraph.descendants(Text('a'))[0]?.attr(Xlink('href')) ?? null;
    if (href !== null) return { target: href };
  }
  return null;
}

/**
 * Read a `table:named-expressions` block: each `table:named-range` the model can hold becomes a
 * named range (sheet-local when `scope` is a sheet name); everything else — a named expression
 * (`Rate = of:=0.07`), a range the model cannot parse — is returned to be kept verbatim.
 */
function readNamedExpressions(container: XmlElement, workbook: Workbook, scope: string | null): XmlElement[] {
  const preserved: XmlElement[] = [];
  for (const named of container.elements()) {
    if (named.is(Table('named-range'))) {
      const name = named.attr(Table('name'));
      const address = named.attr(Table('cell-range-address'));
      if (name !== null && address !== null) {
        const internalRef = OpenFormulaTranslator.fromOpenFormula(
          'of:=[' + address.replace(/^\$+/, '') + ']',
        );
        const range = CellRange.tryParse(internalRef);
        if (range) {
          const key = name.toLowerCase();
          const taken = workbook.namedRanges.some(
            (n) => n.name.toLowerCase() === key && (n.scope ?? null)?.toLowerCase() === scope?.toLowerCase(),
          );
          if (!taken)
            workbook.addNamedRange(scope === null ? { name, target: range } : { name, target: range, scope });
          continue;
        }
      }
    }
    preserved.push(named.clone());
  }
  return preserved;
}

/** `int.TryParse(…) && n > 0 ? n : 1`. */
function readPositive(element: XmlElement, attribute: XmlName): number {
  const text = element.attr(attribute)?.trim() ?? '';
  if (!/^[+-]?\d+$/.test(text)) return 1;
  const n = Number(text);
  return n > 0 && n <= 2147483647 ? n : 1;
}

function parseDouble(text: string | null): number | null {
  switch (text) {
    case 'INF':
      return Infinity;
    case '-INF':
      return -Infinity;
    case 'NaN':
      return NaN;
  }
  return parseInvariantFloat(text);
}

const ISO_DATE =
  /^\s*(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?\s*(?:Z|[+-]\d{2}:?\d{2})?\s*$/;

/** An ODF `office:date-value` (`yyyy-MM-dd` or `yyyy-MM-ddTHH:mm:ss[.fff]`) → a day serial. */
function parseDateValue(text: string | null): number | null {
  if (text === null) return null;
  const m = ISO_DATE.exec(text);
  if (!m) return null;
  const year = Number(m[1]);
  const month = Number(m[2]);
  const day = Number(m[3]);
  const hour = m[4] === undefined ? 0 : Number(m[4]);
  const minute = m[5] === undefined ? 0 : Number(m[5]);
  const second = m[6] === undefined ? 0 : Number(m[6]);
  const millisecond = m[7] === undefined ? 0 : Math.round(Number('0.' + m[7]) * 1000);
  if (month < 1 || month > 12 || day < 1 || day > daysInMonth(year, month)) return null;
  if (hour > 23 || minute > 59 || second > 59) return null;
  return DateConversion.toSerial({ year, month, day, hour, minute, second, millisecond });
}

function daysInMonth(year: number, month: number): number {
  if (month === 2) return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0 ? 29 : 28;
  return [4, 6, 9, 11].includes(month) ? 30 : 31;
}

const DURATION =
  /^\s*(-)?P(?:(\d+)Y)?(?:(\d+)M)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d*)?|\.\d+)S)?)?\s*$/;

/**
 * `XmlConvert.ToTimeSpan(text).TotalDays`: an `xsd:duration` → days, with .NET's conventions of a
 * 365-day year and a 30-day month. Null when malformed (including a bare `P` or a dangling `T`).
 */
export function parseDuration(text: string | null): number | null {
  if (!text) return null;
  const m = DURATION.exec(text);
  if (!m) return null;
  const hasDate = m[2] !== undefined || m[3] !== undefined || m[4] !== undefined;
  const hasTime = m[5] !== undefined || m[6] !== undefined || m[7] !== undefined;
  if (!hasDate && !hasTime) return null;
  if (text.includes('T') && !hasTime) return null;

  const n = (s: string | undefined): number => (s === undefined ? 0 : Number(s));
  const seconds = ((n(m[2]) * 365 + n(m[3]) * 30 + n(m[4])) * 24 + n(m[5])) * 3600 + n(m[6]) * 60 + n(m[7]);
  const ticks = Math.round(seconds * 10_000_000);
  const days = ticks / 864_000_000_000;
  return m[1] ? -days : days;
}
