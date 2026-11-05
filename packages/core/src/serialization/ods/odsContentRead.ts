import { DateConversion } from '../../formulas/evaluation/dateConversion.ts';
import { OpenFormulaTranslator } from '../../formulas/openFormulaTranslator.ts';
import { CellAddress } from '../../model/cellAddress.ts';
import { CellRange } from '../../model/cellRange.ts';
import type { Hyperlink } from '../../model/hyperlink.ts';
import { MergedRegion } from '../../model/mergedRegion.ts';
import { SheetDimensions } from '../../model/sheetDimensions.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { StylePool } from '../../styling/stylePool.ts';
import { CellValue } from '../../values/cellValue.ts';
import { serializeXmlElement, XmlElement, XmlText, type XmlDocument, type XmlName } from '../../xml/index.ts';
import { parseInvariantFloat } from './mapping/odsCellStyles.ts';
import { OdsCharts } from './odsCharts.ts';
import { readDatabaseRanges } from './odsDatabaseRanges.ts';
import type { ForeignXml, OdsPackage } from './odsPackage.ts';
import { OdsPreserved, type PreservedCell } from './odsPreserved.ts';
import { OdsStyleIndex } from './odsStyleIndex.ts';
import {
  assembleValidations,
  readContentValidationRules,
  recordValidationRect,
  type ValidationCells,
} from './odsValidation.ts';
import { Draw, OdsXmlNames as N, Office, Table, Text, Xlink } from './odsXmlNames.ts';

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

  let seenTable = false;
  for (const child of spreadsheet.elements()) {
    if (child.is(Table('table'))) {
      seenTable = true;
      readTable(child, workbook, styles, validationCells, pkg);
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
  OdsPreserved.setWorkbook(workbook, {
    namedExpressions,
    automaticStyles: styles.preservedAutomaticStyles,
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
): void {
  const name = table.attr(Table('name')) ?? workbook.addSheet().name;
  const sheet = workbook.findSheet(name) ?? workbook.addSheetForLoad(name);

  // A Cellestra chart sheet is a table holding nothing but the chart's frame; the chart is modeled
  // and its frame regenerated. An object in a sheet that has data is preserved like any shape.
  const chartHref = OdsCharts.readChartSheet(table, pkg, sheet);
  const context: ReadContext = { styles, validationCells, chartSheet: chartHref !== null };

  readColumns(table, sheet, styles);
  readRows(table, sheet, context);

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
    if (chartHref !== null && OdsCharts.containsObject(child, chartHref)) continue;
    (seenGrid ? after : before).push(child.clone());
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

function readColumns(table: XmlElement, sheet: Worksheet, styles: OdsStyleIndex): void {
  let column = 0;
  for (const element of gridElements(table, Table('table-column'), COLUMN_GROUPS)) {
    if (column >= MAX_COLUMNS) break;
    const repeat = Math.min(readPositive(element, Table('number-columns-repeated')), MAX_COLUMNS - column);
    const styleName = element.attr(Table('style-name'));
    const hidden = isHidden(element);
    const width = styleName === null ? null : styles.columnWidth(styleName);
    // The column's default cell style: "this whole column is formatted like so".
    const defaultStyleId = defaultStyleIdOf(element, styles);

    if (width === null && !hidden && defaultStyleId === null) {
      column += repeat;
      continue;
    }
    for (let i = 0; i < repeat; i++, column++) {
      const model = sheet.getOrCreateColumn(column);
      if (width !== null) model.width = width;
      if (defaultStyleId !== null) model.styleId = defaultStyleId;
      model.hidden = hidden;
    }
  }
}

function readRows(table: XmlElement, sheet: Worksheet, context: ReadContext): void {
  const rows = gridElements(table, Table('table-row'), ROW_GROUPS);
  let row = 0;
  for (let index = 0; index < rows.length; index++) {
    if (row >= MAX_ROWS) break;
    const element = rows[index]!;
    const repeat = Math.min(readPositive(element, Table('number-rows-repeated')), MAX_ROWS - row);
    const styleName = element.attr(Table('style-name'));
    const hidden = isHidden(element);
    const height = styleName === null ? null : context.styles.rowHeight(styleName);
    const defaultStyleId = defaultStyleIdOf(element, context.styles);
    const trailing = index === rows.length - 1 || row + repeat >= MAX_ROWS;
    const inert = rowIsInert(element);

    // LibreOffice pads sheets with runs like 1,048,000 repeated rows of 1,024 blank cells; a plain
    // run carries nothing, so skip it whole instead of visiting a billion positions. The padding
    // run at the end of a sheet only ever carries the default row height, so it is skipped too.
    if (inert && !hidden && defaultStyleId === null && (height === null || trailing)) {
      row += repeat;
      continue;
    }
    if (height !== null || hidden || defaultStyleId !== null) {
      for (let i = 0; i < repeat; i++) {
        const model = sheet.getOrCreateRow(row + i);
        if (height !== null) model.height = height;
        if (defaultStyleId !== null) model.styleId = defaultStyleId;
        model.hidden = hidden;
      }
    }
    if (!inert) readCells(element, sheet, row, repeat, trailing, context);
    row += repeat;
  }
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
  readonly styleId: number | null;
  readonly hyperlink: Hyperlink | null;
  readonly preserved: PreservedCell | null;
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
        for (let c = column; c < column + repeat; c++) sheet.getOrCreateColumn(c).styleId = parsed.styleId!;
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
            sheet.addMerge(new MergedRegion(range));
          }
        }
      }
    }
    column += repeat;
  }
}

function storeCell(sheet: Worksheet, row: number, column: number, parsed: ParsedCell): void {
  const cell = sheet.getOrCreate(row, column);
  cell.value = parsed.value;
  cell.formula = parsed.formula;
  cell.hyperlink = parsed.hyperlink;
  if (parsed.styleId !== null) cell.styleId = parsed.styleId;
  if (parsed.preserved) OdsPreserved.setCell(sheet, row, column, cell, parsed.preserved);
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
    (paragraphs.length > 0 ? after : before).push(child.clone());
  }

  return {
    value: readValue(element, paragraphs),
    formula: formula !== null ? OpenFormulaTranslator.fromOpenFormula(formula) : null,
    styleId: styleName === null ? null : context.styles.cellStyleId(styleName),
    hyperlink: readHyperlink(paragraphs),
    preserved: before.length > 0 || after.length > 0 ? { before, after } : null,
  };
}

function readValue(element: XmlElement, paragraphs: readonly XmlElement[]): CellValue {
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
    case 'string':
      return CellValue.text(paragraphs.map(paragraphText).join('\n'));
    default:
      return CellValue.empty;
  }
}

/** The most spaces one `text:s` may stand for (a hostile `text:c` would otherwise allocate gigabytes). */
const MAX_SPACE_RUN = 65536;

/**
 * The text of one `text:p` (ODF 6.1): `text:s` is `text:c` spaces (default 1), `text:tab` a tab,
 * `text:line-break` a newline, and spans, links and other inline markup contribute their own text.
 * Whitespace-only text nodes are content too (a cell holding " " is a space). Only paragraphs are
 * joined with "\n", by the caller. Iterative, so deeply nested spans cannot overflow the stack.
 */
function paragraphText(paragraph: XmlElement): string {
  let text = '';
  const stack: Array<{ e: XmlElement; i: number }> = [{ e: paragraph, i: 0 }];
  while (stack.length > 0) {
    const top = stack[stack.length - 1]!;
    if (top.i >= top.e.nodes.length) {
      stack.pop();
      continue;
    }
    const node = top.e.nodes[top.i++]!;
    if (node instanceof XmlText) {
      text += node.value;
      continue;
    }
    if (!(node instanceof XmlElement)) continue;
    if (node.uri === Text.uri) {
      if (node.local === 's') {
        const count = readPositive(node, Text('c'));
        text += ' '.repeat(Math.min(count, MAX_SPACE_RUN));
        continue;
      }
      if (node.local === 'tab') {
        text += '\t';
        continue;
      }
      if (node.local === 'line-break') {
        text += '\n';
        continue;
      }
      // A footnote or an annotation inside the paragraph is not part of the cell's text.
      if (node.local === 'note') continue;
    }
    if (node.is(Office('annotation'))) continue;
    stack.push({ e: node, i: 0 });
  }
  return text;
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
