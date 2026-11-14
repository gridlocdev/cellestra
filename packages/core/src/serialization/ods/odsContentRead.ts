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
import { serializeXmlElement, XmlElement, type XmlDocument, type XmlName } from '../../xml/index.ts';
import { parseInvariantFloat } from './mapping/odsCellStyles.ts';
import { OdsCharts } from './odsCharts.ts';
import {
  applyCheckboxes,
  collectCheckboxes,
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
  /** The sheet is a chart sheet, so its chart frame is modeled (not preserved) content. */
  readonly chartSheet: boolean;
  /** Checkbox controls the model takes over; their cell-anchored `draw:control`s are dropped. */
  readonly checkboxIds: ReadonlySet<string>;
  /** The table's merges, gathered in reading order and added once the grid is read. */
  readonly merges: MergedRegion[];
}

/**
 * Parse `content.xml` into `workbook`. Unmodeled `office:spreadsheet` children are appended to
 * `foreign` as verbatim XML (with whether they came before the first table) so the writer can
 * re-emit them in place; unmodeled table, cell and named-expression content is kept through
 * `OdsPreserved`. `stylesDocument` is styles.xml, whose common and data styles cells may name.
 */
export function readContentXml(
  document: XmlDocument,
  workbook: Workbook,
  foreign: ForeignXml[],
  pkg: OdsPackage | null = null,
  stylesDocument: XmlDocument | null = null,
): void {
  const root = document.root;
  const styles = new OdsStyleIndex(root, stylesDocument?.root ?? null, workbook.styles);

  const spreadsheet = root.element(Office('body'))?.element(Office('spreadsheet')) ?? null;
  if (!spreadsheet) return;

  // Rule definitions first; cells reference them by name and each rule's target is rebuilt from
  // the cells that point at it.
  const validationRules = readContentValidationRules(spreadsheet);
  const validationCells: ValidationCells = new Map();
  const namedExpressions: XmlElement[] = [];
  const checkboxes: ReadCheckboxes = { ids: new Set(), cells: [] };

  let seenTable = false;
  for (const child of spreadsheet.elements()) {
    if (child.is(Table('table'))) {
      seenTable = true;
      readTable(child, workbook, styles, validationCells, pkg, checkboxes);
    } else if (child.is(Table('named-expressions'))) {
      namedExpressions.push(...readNamedExpressions(child, workbook, null));
    } else if (child.is(Table(N.databaseRanges))) {
      readDatabaseRanges(child, workbook);
    } else if (child.is(Table(N.contentValidations))) {
      // Modeled; the rules were parsed above. Not preserved as foreign.
    } else {
      foreign.push({ beforeTables: !seenTable, xml: serializeXmlElement(child) });
    }
  }

  assembleValidations(workbook, validationRules, validationCells);
  applyCheckboxes(workbook, checkboxes);
  OdsPreserved.setWorkbook(workbook, {
    namedExpressions,
    automaticStyles: styles.preservedAutomaticStyles,
    fontFaces: root.element(Office('font-face-decls'))?.clone() ?? null,
  });
}

/** The `leaf` elements of a grid in document order, through any nesting of `groups`. Iterative. */
function gridElements(table: XmlElement, leaf: XmlName, groups: readonly XmlName[]): XmlElement[] {
  const out: XmlElement[] = [];
  const stack: Array<{ e: XmlElement; i: number }> = [{ e: table, i: 0 }];
  while (stack.length > 0) {
    const top = stack[stack.length - 1]!;
    if (top.i >= top.e.nodes.length) {
      stack.pop();
      continue;
    }
    const n = top.e.nodes[top.i++]!;
    if (!(n instanceof XmlElement)) continue;
    if (n.is(leaf)) out.push(n);
    else if (groups.some((g) => n.is(g))) stack.push({ e: n, i: 0 });
  }
  return out;
}

const GRID_ELEMENTS: readonly XmlName[] = [
  Table('table-column'),
  Table('table-row'),
  ...ROW_GROUPS,
  ...COLUMN_GROUPS,
];

function readTable(
  table: XmlElement,
  workbook: Workbook,
  styles: OdsStyleIndex,
  validationCells: ValidationCells,
  pkg: OdsPackage | null,
  checkboxes: ReadCheckboxes,
): void {
  const name = table.attr(Table('name')) ?? workbook.addSheet().name;
  const sheet = workbook.findSheet(name) ?? workbook.addSheetForLoad(name);

  // A Cellestra chart sheet is a table holding nothing but the chart's frame in A1; the chart is
  // modeled and its frame regenerated. A chart anywhere else is preserved like any shape.
  const chartHref = OdsCharts.readChartSheet(table, pkg, sheet);
  collectCheckboxes(table, sheet.name, checkboxes);
  const context: ReadContext = {
    styles,
    validationCells,
    chartSheet: chartHref !== null,
    checkboxIds: checkboxes.ids,
    merges: [],
  };

  const grid = new TableGridReader(sheet, context);
  for (const element of gridElements(table, Table('table-column'), COLUMN_GROUPS)) grid.column(element);
  for (const element of gridElements(table, Table('table-row'), ROW_GROUPS)) grid.row(element);
  grid.finish();

  // Everything else the table carries is kept as it was: its attributes (protection, print
  // ranges, its table style) and its unmodeled children (shapes, forms, conditional formats).
  const before: XmlElement[] = [];
  const after: XmlElement[] = [];
  const namedExpressions: XmlElement[] = [];
  let seenGrid = false;
  for (const child of table.elements()) {
    if (GRID_ELEMENTS.some((g) => child.is(g))) {
      seenGrid = true;
      continue;
    }
    if (child.is(Table('named-expressions'))) {
      namedExpressions.push(...readNamedExpressions(child, workbook, sheet.name));
      continue;
    }
    // The checkboxes the model took over are regenerated on save; the rest of the forms stay.
    const kept = withoutCheckboxes(child, checkboxes.ids);
    if (kept) (seenGrid ? after : before).push(kept === child ? child.clone() : kept);
  }
  const attributes = table.attributes.filter((a) => !(a.name.uri === Table.uri && a.name.local === 'name'));
  if (attributes.length > 0 || before.length > 0 || after.length > 0 || namedExpressions.length > 0) {
    OdsPreserved.setTable(sheet, {
      attributes: attributes.map((a) => ({ name: a.name, value: a.value })),
      before,
      after,
      namedExpressions,
    });
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
  for (let index = 0; index < cells.length; index++) {
    if (column >= MAX_COLUMNS) break;
    const element = cells[index]!;
    const repeat = Math.min(readPositive(element, Table('number-columns-repeated')), MAX_COLUMNS - column);
    if (cellIsInert(element)) {
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
      if (
        !hasContent &&
        lastInRow &&
        (area > STYLED_RUN_CELL_LIMIT || (column + repeat >= MAX_COLUMNS && repeat > 1))
      ) {
        // A styled run to the end of the row (LibreOffice writes one for a formatted whole row):
        // that is a row format, not 16384 cells.
        for (let r = row; r < row + rowCount; r++) sheet.getOrCreateRow(r).styleId = parsed.styleId!;
      } else if (!hasContent && area > STYLED_RUN_CELL_LIMIT && trailingRows) {
        // A styled run down to the end of the sheet: that is a column format.
        columnFormats.push({ from: column, to: column + repeat, styleId: parsed.styleId! });
      } else {
        if (area > CONTENT_RUN_CELL_LIMIT) {
          throw new OdsContentTooLargeError(
            `Sheet '${sheet.name}' repeats one cell over ${area} positions; the most a run may cover is ${CONTENT_RUN_CELL_LIMIT}.`,
          );
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
      if (columnsSpanned > 1 || readPositive(element, Table('number-rows-spanned')) > 1) {
        for (let r = row; r < row + rowCount; r++) {
          const rowsSpanned = Math.min(readPositive(element, Table('number-rows-spanned')), MAX_ROWS - r);
          if (columnsSpanned <= 1 && rowsSpanned <= 1) continue;
          for (let c = column; c < column + repeat; c++) {
            const right = Math.min(c + columnsSpanned - 1, MAX_COLUMNS - 1);
            const range = new CellRange(new CellAddress(c, r), new CellAddress(right, r + rowsSpanned - 1));
            context.merges.push(new MergedRegion(range));
          }
        }
      }
    }
    column += repeat;
  }
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
