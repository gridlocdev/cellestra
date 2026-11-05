import { Defaults } from '../../defaults.ts';
import { DateConversion, type CivilDateTime } from '../../formulas/evaluation/dateConversion.ts';
import { OpenFormulaTranslator } from '../../formulas/openFormulaTranslator.ts';
import type { Cell } from '../../model/cell.ts';
import type { MergedRegion } from '../../model/mergedRegion.ts';
import { cellKey } from '../../model/sheetDimensions.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { NumberFormats } from '../../styling/cellStyle.ts';
import { NumberFormatter } from '../../styling/numberFormatter.ts';
import { StylePool } from '../../styling/stylePool.ts';
import { roundTripNumber } from '../../values/cellValue.ts';
import { parseXmlElement, XmlDocument, XmlElement, XmlText } from '../../xml/index.ts';
import { OdsCharts } from './odsCharts.ts';
import { odfCellAddress, odfRangeAddress, type ValidationNames } from './odsContentShared.ts';
import { addDatabaseRanges } from './odsDatabaseRanges.ts';
import type { ForeignXml } from './odsPackage.ts';
import { StyleRegistry } from './odsStyleRegistry.ts';
import { addContentValidations } from './odsValidation.ts';
import {
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

/** C# `CultureInfo.InvariantCulture` for the cached display text written beside each value. */
const INVARIANT_LOCALE = 'en-US';

type CellValidation = ReadonlyMap<number, string>;
const EMPTY_VALIDATION_MAP: CellValidation = new Map();

/**
 * Build `content.xml`: the cell grid, automatic styles, merges, column widths/row heights, hidden
 * lines, named ranges, validations, and database ranges. Unmodeled `office:spreadsheet` children
 * captured on read are re-emitted verbatim in their original position relative to the tables.
 */
export function writeContentXml(workbook: Workbook, foreign: readonly ForeignXml[]): XmlDocument {
  const autoStyles = new XmlElement(Office('automatic-styles'));
  const registry = new StyleRegistry(workbook.styles, autoStyles);

  const spreadsheet = new XmlElement(Office('spreadsheet'));

  for (const child of foreign) if (child.beforeTables) spreadsheet.add(parseXmlElement(child.xml));

  // Data validations are declared once, before the tables (ODF schema order), and referenced by
  // name from each validated cell.
  const validationNames = addContentValidations(spreadsheet, workbook);

  for (const sheet of workbook.worksheets) spreadsheet.add(writeTable(sheet, registry, validationNames));

  addNamedExpressions(spreadsheet, workbook);
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
    .declare(Xlink.prefix, Xlink.uri)
    .declare(Of.prefix, Of.uri)
    .setAttr(Office('version'), '1.3')
    .add(autoStyles, new XmlElement(Office('body'), spreadsheet));

  return new XmlDocument(root);
}

function writeTable(sheet: Worksheet, registry: StyleRegistry, validationNames: ValidationNames): XmlElement {
  const table = new XmlElement(Table('table')).setAttr(Table('name'), sheet.name);
  const used = sheet.getUsedRange();
  let lastColumn = used?.right ?? 0;
  let lastRow = used?.bottom ?? 0;

  // A validated cell must be emitted (with its content-validation-name) even when it is blank and
  // lies past the used range, so extend the written bounds to cover every validation.
  const cellValidation = validationNames.get(sheet.name) ?? EMPTY_VALIDATION_MAP;
  for (const validation of sheet.validations) {
    lastColumn = Math.max(lastColumn, validation.target.right);
    lastRow = Math.max(lastRow, validation.target.bottom);
  }

  // A column that carries formatting but no cells is past the used range and still has to be
  // written, or the formatting is lost on save.
  for (const [index, model] of sheet.columns) if (!model.isDefault) lastColumn = Math.max(lastColumn, index);

  writeColumns(table, sheet, lastColumn, registry);
  writeRows(table, sheet, lastRow, lastColumn, registry, cellValidation);

  // A chart sheet embeds its chart as an ODF draw frame (the object was written by writeObjects).
  if (sheet.isChartSheet) OdsCharts.injectFrame(table, sheet);

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

function writeRows(
  table: XmlElement,
  sheet: Worksheet,
  lastRow: number,
  lastColumn: number,
  registry: StyleRegistry,
  cellValidation: CellValidation,
): void {
  let row = 0;
  while (row <= lastRow) {
    if (rowIsEmpty(sheet, row, lastColumn, cellValidation)) {
      let repeat = 1;
      while (row + repeat <= lastRow && rowIsEmpty(sheet, row + repeat, lastColumn, cellValidation)) repeat++;

      const emptyRow = new XmlElement(Table('table-row'), new XmlElement(Table('table-cell')));
      if (repeat > 1) emptyRow.setAttr(Table('number-rows-repeated'), repeat);
      table.add(emptyRow);
      row += repeat;
      continue;
    }

    table.add(writeRow(sheet, row, lastColumn, registry, cellValidation));
    row++;
  }
}

function rowIsEmpty(
  sheet: Worksheet,
  row: number,
  lastColumn: number,
  cellValidation: CellValidation,
): boolean {
  const model = sheet.findRow(row);
  if (model && !model.isDefault) return false;
  for (let c = 0; c <= lastColumn; c++) {
    if (sheet.find(row, c) || sheet.findMerge(row, c) || cellValidation.has(cellKey(row, c))) return false;
  }
  return true;
}

function writeRow(
  sheet: Worksheet,
  row: number,
  lastColumn: number,
  registry: StyleRegistry,
  cellValidation: CellValidation,
): XmlElement {
  const element = new XmlElement(Table('table-row'));
  const rowModel = sheet.findRow(row);
  if (rowModel && rowModel.height !== Defaults.rowHeight) {
    element.setAttr(Table('style-name'), registry.rowStyle(rowModel.height));
  }
  if (rowModel?.hidden) element.setAttr(Table('visibility'), 'collapse');
  // A whole formatted row travels the same way a formatted column does — one attribute.
  if (rowModel && rowModel.styleId !== StylePool.defaultId) {
    element.setAttr(Table('default-cell-style-name'), registry.cellStyle(rowModel.styleId));
  }

  let column = 0;
  while (column <= lastColumn) {
    const merge = sheet.findMerge(row, column);
    if (merge && merge.isCovered(row, column)) {
      element.add(new XmlElement(Table('covered-table-cell')));
      column++;
      continue;
    }

    const cell = sheet.find(row, column);
    if (!cell && !merge && !cellValidation.has(cellKey(row, column))) {
      // Coalesce a run of blank, unvalidated cells.
      let repeat = 1;
      while (
        column + repeat <= lastColumn &&
        !sheet.find(row, column + repeat) &&
        !sheet.findMerge(row, column + repeat) &&
        !cellValidation.has(cellKey(row, column + repeat))
      ) {
        repeat++;
      }

      const blank = new XmlElement(Table('table-cell'));
      if (repeat > 1) blank.setAttr(Table('number-columns-repeated'), repeat);
      element.add(blank);
      column += repeat;
      continue;
    }

    element.add(writeCell(cell ?? null, merge ?? null, row, column, registry, cellValidation));
    column++;
  }
  return element;
}

function writeCell(
  cell: Cell | null,
  merge: MergedRegion | null,
  row: number,
  column: number,
  registry: StyleRegistry,
  cellValidation: CellValidation,
): XmlElement {
  const element = new XmlElement(Table('table-cell'));

  if (cell && cell.styleId !== StylePool.defaultId) {
    element.setAttr(Table('style-name'), registry.cellStyle(cell.styleId));
  }

  const validationName = cellValidation.get(cellKey(row, column));
  if (validationName !== undefined) element.setAttr(Table(N.contentValidationName), validationName);

  if (merge && merge.range.top === row && merge.range.left === column) {
    if (merge.range.columnCount > 1)
      element.setAttr(Table('number-columns-spanned'), merge.range.columnCount);
    if (merge.range.rowCount > 1) element.setAttr(Table('number-rows-spanned'), merge.range.rowCount);
  }

  if (cell) writeCellContent(element, cell, registry);
  return element;
}

function writeCellContent(element: XmlElement, cell: Cell, registry: StyleRegistry): void {
  if (cell.formula !== null) {
    element.setAttr(Table('formula'), OpenFormulaTranslator.toOpenFormula(cell.formula));
  }

  const value = cell.value;
  const format = registry.styles.get(cell.styleId).numberFormat;
  const kind = NumberFormats.kind(format);
  const isTime = kind === 'time';
  const isDate = kind === 'date' || kind === 'dateTime';

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
      } else {
        element.setAttr(Office('value-type'), 'float');
        element.setAttr(Office('value'), floatValueText(value.value));
      }
      element.add(new XmlElement(Text('p'), NumberFormatter.format(value, format.code, INVARIANT_LOCALE)));
      break;
    case 'boolean':
      element.setAttr(Office('value-type'), 'boolean');
      element.setAttr(Office('boolean-value'), value.value ? 'true' : 'false');
      element.add(new XmlElement(Text('p'), value.value ? 'TRUE' : 'FALSE'));
      break;
    case 'text':
      element.setAttr(Office('value-type'), 'string');
      // One text:p per line: ODF has no newline inside a paragraph (a conformant reader collapses
      // it to a space). The reader joins paragraphs with "\n".
      for (const line of value.value.split('\n')) element.add(new XmlElement(Text('p'), line));
      break;
    case 'error':
      element.setAttr(Office('value-type'), 'string');
      element.add(new XmlElement(Text('p'), value.error.token));
      break;
    case 'empty':
      break;
  }

  const paragraph = element.element(Text('p'));
  if (cell.hyperlink && paragraph) {
    const anchor = new XmlElement(Text('a'), paragraph.value).setAttr(Xlink('href'), cell.hyperlink.target);
    paragraph.value = '';
    // An empty text node marks the paragraph as mixed content so the writer does not indent it;
    // indentation inside text:p is real cell text to a reader.
    paragraph.add(new XmlText(''), anchor);
  }
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
  return midnight
    ? datePart(when)
    : `${datePart(when)}T${pad(when.hour, 2)}:${pad(when.minute, 2)}:${pad(when.second, 2)}`;
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

function addNamedExpressions(spreadsheet: XmlElement, workbook: Workbook): void {
  const global = workbook.namedRanges.filter((n) => n.scope === null || n.scope === undefined);
  if (global.length === 0) return;

  const fallbackSheet =
    workbook.worksheets.length > 0 ? workbook.worksheets[0]!.name : Defaults.firstSheetName;
  const container = new XmlElement(Table('named-expressions'));
  for (const named of global) {
    container.add(
      new XmlElement(Table('named-range'))
        .setAttr(Table('name'), named.name)
        .setAttr(Table('cell-range-address'), odfRangeAddress(named.target, fallbackSheet))
        .setAttr(Table('base-cell-address'), odfCellAddress(named.target.topLeft, fallbackSheet)),
    );
  }
  spreadsheet.add(container);
}
