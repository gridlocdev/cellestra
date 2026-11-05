import { DateConversion } from '../../formulas/evaluation/dateConversion.ts';
import { OpenFormulaTranslator } from '../../formulas/openFormulaTranslator.ts';
import { CellAddress } from '../../model/cellAddress.ts';
import { CellRange } from '../../model/cellRange.ts';
import type { Hyperlink } from '../../model/hyperlink.ts';
import { MergedRegion } from '../../model/mergedRegion.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { NumberFormats, type NumberFormat } from '../../styling/cellStyle.ts';
import type { StylePool } from '../../styling/stylePool.ts';
import { CellValue } from '../../values/cellValue.ts';
import { serializeXmlElement, XmlElement, XmlText, type XmlDocument, type XmlName } from '../../xml/index.ts';
import { OdsCellStyles, parseInvariantFloat } from './mapping/odsCellStyles.ts';
import { OdsDataStyles } from './mapping/odsDataStyles.ts';
import { OdsCharts } from './odsCharts.ts';
import { CM_PER_INCH, cmToPx } from './odsContentShared.ts';
import { readDatabaseRanges } from './odsDatabaseRanges.ts';
import type { ForeignXml, OdsPackage } from './odsPackage.ts';
import {
  assembleValidations,
  readContentValidationRules,
  recordValidationCell,
  type ValidationCells,
} from './odsValidation.ts';
import { NumberNs, OdsXmlNames as N, Office, Style, Table, Text, Xlink } from './odsXmlNames.ts';

interface AutomaticStyles {
  readonly styleIds: ReadonlyMap<string, number>;
  readonly columnWidths: ReadonlyMap<string, number>;
  readonly rowHeights: ReadonlyMap<string, number>;
}

/**
 * Parse `content.xml` into `workbook`. Unmodeled `office:spreadsheet` children are appended to
 * `foreign` as verbatim XML (with whether they came before the first table) so the writer can
 * re-emit them in place.
 */
export function readContentXml(
  document: XmlDocument,
  workbook: Workbook,
  foreign: ForeignXml[],
  pkg: OdsPackage | null = null,
): void {
  const root = document.root;
  const styles = readAutomaticStyles(root, workbook.styles);

  const spreadsheet = root.element(Office('body'))?.element(Office('spreadsheet')) ?? null;
  if (!spreadsheet) return;

  // Rule definitions first; cells reference them by name and each rule's target is rebuilt from
  // the cells that point at it.
  const validationRules = readContentValidationRules(spreadsheet);
  const validationCells: ValidationCells = new Map();

  let seenTable = false;
  for (const child of spreadsheet.elements()) {
    if (child.is(Table('table'))) {
      seenTable = true;
      readTable(child, workbook, styles, validationCells, pkg);
    } else if (child.is(Table('named-expressions'))) {
      readNamedExpressions(child, workbook);
    } else if (child.is(Table(N.databaseRanges))) {
      readDatabaseRanges(child, workbook);
    } else if (child.is(Table(N.contentValidations))) {
      // Modeled; the rules were parsed above. Not preserved as foreign.
    } else {
      foreign.push({ beforeTables: !seenTable, xml: serializeXmlElement(child) });
    }
  }

  assembleValidations(workbook, validationRules, validationCells);
}

function readAutomaticStyles(root: XmlElement, pool: StylePool): AutomaticStyles {
  const styleIds = new Map<string, number>();
  const columnWidths = new Map<string, number>();
  const rowHeights = new Map<string, number>();

  const autoStyles = root.element(Office('automatic-styles'));
  if (!autoStyles) return { styleIds, columnWidths, rowHeights };

  const dataStyles = new Map<string, NumberFormat>();
  for (const element of autoStyles.elements()) {
    if (element.uri === NumberNs.uri) {
      const name = element.attr(Style('name'));
      if (name !== null) dataStyles.set(name, NumberFormats.of(OdsDataStyles.toCode(element)));
    }
  }

  for (const element of autoStyles.elements(Style('style'))) {
    const name = element.attr(Style('name'));
    if (name === null) continue;

    switch (element.attr(Style('family'))) {
      case 'table-cell': {
        const cellStyle = OdsCellStyles.parse(
          element,
          (code) => dataStyles.get(code) ?? NumberFormats.general,
        );
        styleIds.set(name, pool.intern(cellStyle));
        break;
      }
      case 'table-column': {
        const width = parseCm(
          element.element(Style('table-column-properties'))?.attr(Style('column-width')) ?? null,
        );
        if (width !== null) columnWidths.set(name, cmToPx(width));
        break;
      }
      case 'table-row': {
        const height = parseCm(
          element.element(Style('table-row-properties'))?.attr(Style('row-height')) ?? null,
        );
        if (height !== null) rowHeights.set(name, cmToPx(height));
        break;
      }
    }
  }

  return { styleIds, columnWidths, rowHeights };
}

function readTable(
  table: XmlElement,
  workbook: Workbook,
  styles: AutomaticStyles,
  validationCells: ValidationCells,
  pkg: OdsPackage | null,
): void {
  const name = table.attr(Table('name')) ?? workbook.addSheet().name;
  const sheet = workbook.findSheet(name) ?? workbook.addSheetForLoad(name);

  readColumns(table, sheet, styles);
  readRows(table, sheet, styles, validationCells);

  // If the table embeds a chart object, attach it so the sheet opens as a chart sheet.
  OdsCharts.readChart(table, pkg, sheet);
}

function isHidden(element: XmlElement): boolean {
  const visibility = element.attr(Table('visibility'));
  return visibility === 'collapse' || visibility === 'filter';
}

function defaultStyleIdOf(element: XmlElement, styles: AutomaticStyles): number | null {
  const name = element.attr(Table('default-cell-style-name'));
  return name === null ? null : (styles.styleIds.get(name) ?? null);
}

function readColumns(table: XmlElement, sheet: Worksheet, styles: AutomaticStyles): void {
  let column = 0;
  for (const element of table.elements(Table('table-column'))) {
    const repeat = readPositive(element, Table('number-columns-repeated'));
    const styleName = element.attr(Table('style-name'));
    const hidden = isHidden(element);
    const width = styleName === null ? null : (styles.columnWidths.get(styleName) ?? null);
    // The column's default cell style: "this whole column is formatted like so".
    const defaultStyleId = defaultStyleIdOf(element, styles);

    for (let i = 0; i < repeat; i++, column++) {
      if (width === null && !hidden && defaultStyleId === null) continue;
      const model = sheet.getOrCreateColumn(column);
      if (width !== null) model.width = width;
      if (defaultStyleId !== null) model.styleId = defaultStyleId;
      model.hidden = hidden;
    }
  }
}

function readRows(
  table: XmlElement,
  sheet: Worksheet,
  styles: AutomaticStyles,
  validationCells: ValidationCells,
): void {
  let row = 0;
  for (const element of table.elements(Table('table-row'))) {
    const repeat = readPositive(element, Table('number-rows-repeated'));
    const styleName = element.attr(Table('style-name'));
    const hidden = isHidden(element);
    const height = styleName === null ? null : (styles.rowHeights.get(styleName) ?? null);
    const defaultStyleId = defaultStyleIdOf(element, styles);

    for (let i = 0; i < repeat; i++, row++) {
      if (height !== null || hidden || defaultStyleId !== null) {
        const model = sheet.getOrCreateRow(row);
        if (height !== null) model.height = height;
        if (defaultStyleId !== null) model.styleId = defaultStyleId;
        model.hidden = hidden;
      }
      readCells(element, sheet, row, styles, validationCells);
    }
  }
}

function readCells(
  rowElement: XmlElement,
  sheet: Worksheet,
  row: number,
  styles: AutomaticStyles,
  validationCells: ValidationCells,
): void {
  let column = 0;
  for (const element of rowElement.elements()) {
    const covered = element.is(Table('covered-table-cell'));
    if (!covered && !element.is(Table('table-cell'))) continue;

    const repeat = readPositive(element, Table('number-columns-repeated'));
    for (let i = 0; i < repeat; i++, column++) {
      if (covered) continue;
      readCell(element, sheet, row, column, styles, validationCells);
    }
  }
}

function readCell(
  element: XmlElement,
  sheet: Worksheet,
  row: number,
  column: number,
  styles: AutomaticStyles,
  validationCells: ValidationCells,
): void {
  const value = readValue(element);
  const formula = element.attr(Table('formula'));
  const styleName = element.attr(Table('style-name'));
  const hyperlink = readHyperlink(element);

  const validationName = element.attr(Table(N.contentValidationName));
  if (validationName !== null) recordValidationCell(validationCells, sheet.name, validationName, row, column);

  const columnsSpanned = readPositive(element, Table('number-columns-spanned'));
  const rowsSpanned = readPositive(element, Table('number-rows-spanned'));

  const hasContent = value.kind !== 'empty' || formula !== null || styleName !== null || hyperlink !== null;
  if (hasContent) {
    const cell = sheet.getOrCreate(row, column);
    cell.value = value;
    cell.formula = formula !== null ? OpenFormulaTranslator.fromOpenFormula(formula) : null;
    cell.hyperlink = hyperlink;
    const styleId = styleName === null ? undefined : styles.styleIds.get(styleName);
    if (styleId !== undefined) cell.styleId = styleId;
  }

  if (columnsSpanned > 1 || rowsSpanned > 1) {
    const range = new CellRange(
      new CellAddress(column, row),
      new CellAddress(column + columnsSpanned - 1, row + rowsSpanned - 1),
    );
    sheet.addMerge(new MergedRegion(range));
  }
}

function readValue(element: XmlElement): CellValue {
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
    case 'boolean':
      return CellValue.boolean(element.attr(Office('boolean-value')) === 'true');
    case 'string':
      return CellValue.text(readText(element));
    default:
      return CellValue.empty;
  }
}

function readText(element: XmlElement): string {
  return element.elements(Text('p')).map(cellParagraphText).join('\n');
}

/**
 * The text of one `text:p`. A plain paragraph's value is significant as-is. A paragraph with child
 * spans (a hyperlink `text:a`) is mixed content: whitespace-only text nodes around the spans are
 * pretty-print indentation, not content, so they are dropped.
 */
function cellParagraphText(paragraph: XmlElement): string {
  if (!paragraph.hasElements) return paragraph.value;
  let text = '';
  for (const node of paragraph.nodes) {
    if (node instanceof XmlElement) text += node.value;
    else if (node instanceof XmlText && node.value.trim().length > 0) text += node.value;
  }
  return text;
}

function readHyperlink(element: XmlElement): Hyperlink | null {
  const anchor = element.descendants(Text('a'))[0];
  const href = anchor?.attr(Xlink('href')) ?? null;
  return href === null ? null : { target: href };
}

function readNamedExpressions(container: XmlElement, workbook: Workbook): void {
  for (const named of container.elements(Table('named-range'))) {
    const name = named.attr(Table('name'));
    const address = named.attr(Table('cell-range-address'));
    if (name === null || address === null) continue;

    const internalRef = OpenFormulaTranslator.fromOpenFormula('of:=[' + address.replace(/^\$+/, '') + ']');
    const range = CellRange.tryParse(internalRef);
    if (range && !workbook.findNamedRange(name)) workbook.addNamedRange({ name, target: range });
  }
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

function parseCm(value: string | null): number | null {
  if (value === null) return null;
  const lower = value.toLowerCase();
  if (lower.endsWith('cm')) return parseInvariantFloat(value.slice(0, -2));
  if (lower.endsWith('in')) {
    const inches = parseInvariantFloat(value.slice(0, -2));
    return inches === null ? null : inches * CM_PER_INCH;
  }
  return null;
}
