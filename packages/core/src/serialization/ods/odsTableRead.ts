import { OpenFormulaTranslator } from '../../formulas/openFormulaTranslator.ts';
import type { CellRichText, UnmodeledCellContent, UnparsedFormulaSource } from '../../model/cell.ts';
import type { Hyperlink } from '../../model/hyperlink.ts';
import type { MergedRegion } from '../../model/mergedRegion.ts';
import { SheetOutline, type LineGroup, type LineSpan } from '../../model/sheetOutline.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { StylePool } from '../../styling/stylePool.ts';
import type { CellValue } from '../../values/cellValue.ts';
import type { XmlElement } from '../../xml/index.ts';
import { readCellValue } from './odsCellValue.ts';
import { cellStyleIdOf, pinCells, pinsAbove, runFormat, type ColumnFormat } from './odsCellStyleRuns.ts';
import { ChartSheetProbe, OdsCharts } from './odsCharts.ts';
import {
  CheckboxCollector,
  isTakenControl,
  withoutCheckboxes,
  type ReadCheckboxes,
} from './odsCheckboxes.ts';
import { paragraphText } from './odsContentShared.ts';
import { readColumnLayout, readRowLayout, rowCarriesLayout, rowElementLayout } from './odsLines.ts';
import { addMerges, readMergeSpans } from './odsMerges.ts';
import { readNamedExpressions } from './odsNamedExpressions.ts';
import type { OdsPackage } from './odsPackage.ts';
import { keptTableAttributes, OdsPreserved, splitCellChildren } from './odsPreserved.ts';
import {
  chargeRepeatedCells,
  CONTENT_RUN_CELL_LIMIT,
  MAX_COLUMNS,
  MAX_ROWS,
  OdsContentTooLargeError,
  readPositive,
  type RepeatBudget,
} from './odsReadLimits.ts';
import type { OdsStyleIndex } from './odsStyleIndex.ts';
import { recordValidationRect, type ValidationCells } from './odsValidation.ts';
import { Draw, OdsXmlNames as N, Office, Table, Text, Xlink } from './odsXmlNames.ts';

/**
 * Reading one `table:table`: its grid one row element at a time, through the feature modules
 * (`odsLines`, `odsHiddenLines`, `odsMerges`, `odsCellStyleRuns`, `odsCheckboxes`), and its other
 * children and attributes, kept through `odsPreserved`.
 */

export interface ReadContext {
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

/** A row or column container being read: where it starts, and what it says about its lines. */
export interface OpenSpan {
  readonly header: boolean;
  readonly first: number;
  readonly collapsed: boolean;
}

/**
 * Reads one `table:table` as its children arrive. Grid rows go to a `TableGridReader`, except
 * while the table may still be a chart sheet or a checkbox still waits for its linked cell: those
 * change how a cell is read, so rows are held until they are settled (a chart sheet has no data,
 * and a checkbox's control sits in its own cell, so little is ever held). Every other child of the
 * table is kept, and read when the table ends.
 */
export class TableReader {
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
  /** The next row and column a grid element starts at, counted as they arrive (held rows too). */
  private nextRow = 0;
  private nextColumn = 0;
  private readonly rowGroups: LineGroup[] = [];
  private readonly columnGroups: LineGroup[] = [];
  private headerRows: LineSpan | null = null;
  private headerColumns: LineSpan | null = null;

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
    this.nextColumn = Math.min(
      this.nextColumn + readPositive(element, Table('number-columns-repeated')),
      MAX_COLUMNS,
    );
    this.grid.column(element);
  }

  /** Where the next row or column will start. */
  position(kind: 'rows' | 'columns'): number {
    return kind === 'rows' ? this.nextRow : this.nextColumn;
  }

  /** A group or print-title container closed: it covered its first line up to the current one. */
  closeSpan(kind: 'rows' | 'columns', span: OpenSpan): void {
    const last = this.position(kind) - 1;
    if (last < span.first) return;
    if (span.header) {
      // ODF allows one set of print titles per axis; a second is ignored, as LibreOffice ignores it.
      if (kind === 'rows') this.headerRows ??= { first: span.first, last };
      else this.headerColumns ??= { first: span.first, last };
      return;
    }
    const group = { first: span.first, last, collapsed: span.collapsed };
    (kind === 'rows' ? this.rowGroups : this.columnGroups).push(group);
  }

  row(element: XmlElement): void {
    this.nextRow = Math.min(this.nextRow + readPositive(element, Table('number-rows-repeated')), MAX_ROWS);
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
    const byNesting = (a: LineSpan, b: LineSpan): number => a.first - b.first || b.last - a.last;
    const outline: SheetOutline = {
      rowGroups: this.rowGroups.sort(byNesting),
      columnGroups: this.columnGroups.sort(byNesting),
      headerRows: this.headerRows,
      headerColumns: this.headerColumns,
    };
    if (!SheetOutline.isEmpty(outline)) this.sheet.outline = outline;

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
    const attributes = keptTableAttributes(this.outline);
    if (attributes.length > 0 || before.length > 0 || after.length > 0 || namedExpressions.length > 0) {
      OdsPreserved.setTable(this.sheet, {
        attributes,
        before,
        after,
        namedExpressions,
      });
    }
  }
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
    readColumnLayout(this.sheet, element, this.context.styles, column, repeat);
  }

  row(element: XmlElement): void {
    // Rows past the last one are not read, but still count when deciding what the padding is.
    const row = Math.min(this.nextRow, MAX_ROWS);
    const repeat = Math.min(readPositive(element, Table('number-rows-repeated')), MAX_ROWS - row);
    this.nextRow = row + repeat;
    const inert = rowIsInert(element);
    const layout = inert ? rowElementLayout(element) : '';
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
    readRowLayout(this.sheet, element, this.context.styles, row, row + repeat);
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
    if (!rowCarriesLayout(element, this.context.styles, skipHeight)) return;
    readRowLayout(this.sheet, element, this.context.styles, from, to);
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
      const format = runFormat({
        column,
        repeat,
        rowCount,
        hasContent,
        styleId: parsed.styleId,
        lastInRow: index === cells.length - 1,
        rowStyleBefore,
        trailingRows,
      });
      const tooBigForCells = area > CONTENT_RUN_CELL_LIMIT;
      // Columns formatted below a header: the cells above that a column format would restyle.
      const pins =
        format === 'column' ? pinsAbove(sheet, row, column, repeat, parsed.styleId!, context.styles) : null;
      if (format === 'row') {
        // A styled run to the end of the row (LibreOffice writes one for a formatted whole row):
        // that is a row format, not 16384 cells.
        for (let r = row; r < row + rowCount; r++) sheet.getOrCreateRow(r).styleId = parsed.styleId!;
      } else if (format === 'column' && (pins !== null || tooBigForCells)) {
        // A styled run down to the end of the sheet: that is a column format. The few cells above
        // it that look otherwise keep their style; when there are too many to keep, a run too big
        // to hold as cells is a format anyway.
        if (pins !== null) pinCells(sheet, pins, context.repeats);
        columnFormats.push({ from: column, to: column + repeat, styleId: parsed.styleId! });
      } else {
        if (tooBigForCells) {
          throw new OdsContentTooLargeError(
            `Sheet '${sheet.name}' repeats one cell over ${area} positions; the most a run may cover is ${CONTENT_RUN_CELL_LIMIT}.`,
          );
        }
        if (area > 1) chargeRepeatedCells(context.repeats, area, sheet.name);
        for (let r = row; r < row + rowCount; r++) {
          for (let c = column; c < column + repeat; c++) storeCell(sheet, r, c, parsed);
        }
      }
    }

    if (element.is(Table('table-cell'))) {
      readMergeSpans(element, sheet.name, row, rowCount, column, repeat, context.merges, context.repeats);
    }
    passStyle(parsed.styleId ?? StylePool.defaultId);
    column += repeat;
  }
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

  const { paragraphs, preserved } = splitCellChildren(
    element,
    (child) => (context.chartSheet && child.is(Draw('frame'))) || isTakenControl(child, context.checkboxIds),
  );

  const value = readCellValue(element, paragraphs);
  const hyperlink = readHyperlink(paragraphs);
  const read = formula !== null ? OpenFormulaTranslator.read(formula) : null;
  return {
    value,
    formula: read?.formula ?? null,
    formulaSource: read?.source ?? null,
    styleId: cellStyleIdOf(element, context.styles),
    hyperlink,
    preserved,
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

function readHyperlink(paragraphs: readonly XmlElement[]): Hyperlink | null {
  for (const paragraph of paragraphs) {
    const href = paragraph.descendants(Text('a'))[0]?.attr(Xlink('href')) ?? null;
    if (href !== null) return { target: href };
  }
  return null;
}
