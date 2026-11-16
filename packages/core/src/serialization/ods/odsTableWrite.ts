import { Defaults } from '../../defaults.ts';
import { OpenFormulaTranslator } from '../../formulas/openFormulaTranslator.ts';
import type { Cell } from '../../model/cell.ts';
import type { MergedRegion } from '../../model/mergedRegion.ts';
import { CellRange } from '../../model/cellRange.ts';
import { cellKey, SheetDimensions } from '../../model/sheetDimensions.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { XmlElement, XmlText } from '../../xml/index.ts';
import { writeCellStyle } from './odsCellStyleRuns.ts';
import { writeCellValue } from './odsCellValue.ts';
import { OdsCharts } from './odsCharts.ts';
import { addTableForms, checkboxCells, checkboxControl, checkboxForm } from './odsCheckboxes.ts';
import { appendParagraphText } from './odsContentShared.ts';
import { LineContainers } from './odsLineContainers.ts';
import { attributeSignature, RowAppender, rowModelLayout, writeColumns, writeRowLayout } from './odsLines.ts';
import { isSpanned, MergeSweep, writeMergeSpans } from './odsMerges.ts';
import { addLocalNamedExpressions } from './odsNamedExpressions.ts';
import type { PagePlan } from './odsPageSetup.ts';
import { OdsPreserved, writeCellChildren, writeTableAttributes } from './odsPreserved.ts';
import type { StyleRegistry } from './odsStyleRegistry.ts';
import { OdsXmlNames as N, Office, Table, Text, Xlink } from './odsXmlNames.ts';

/**
 * Writing one `table:table`: its kept attributes and children, its checkbox form, then its grid
 * (columns, then rows of cells) through the feature modules: `odsLines` (widths, heights, line
 * formats), `odsHiddenLines`, `odsMerges`, `odsCellStyleRuns`, `odsCheckboxes` and `odsPreserved`.
 */

/** The validations on a sheet, each with the name its cells refer to it by. */
export type ValidationTargets = ReadonlyArray<{ readonly range: CellRange; readonly name: string }>;

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

export function writeTable(
  workbook: Workbook,
  sheet: Worksheet,
  sheetIndex: number,
  registry: StyleRegistry,
  validations: ValidationTargets,
  pages: PagePlan,
): XmlElement {
  const table = new XmlElement(Table('table')).setAttr(Table('name'), sheet.name);
  const preserved = OdsPreserved.table(sheet);
  writeTableAttributes(table, preserved);
  const pageStyle = pages.tableStyleNames.get(sheet);
  if (pageStyle !== undefined) table.setAttr(Table('style-name'), pageStyle);

  // Checkbox cells save as form controls, in the table's `office:forms`.
  const checkboxes = checkboxCells(workbook, sheet, sheetIndex);
  addTableForms(table, preserved?.before ?? [], checkboxes.size > 0 ? checkboxForm(sheet, checkboxes) : null);

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
  if (needsRowAfter(sheet, lastRow, lastColumn)) lastRow++;

  writeColumns(columnContainers, sheet, lastColumn, registry);
  const merges = new MergeSweep(sheet.merges);
  writeRows(rowContainers, { sheet, registry, validations, lastColumn, checkboxes, merges }, lastRow);

  // A chart sheet embeds its chart as an ODF draw frame (the object was written by writeObjects).
  if (sheet.isChartSheet) OdsCharts.injectFrame(table, sheet);

  addLocalNamedExpressions(table, workbook, sheet, preserved?.namedExpressions ?? []);
  for (const child of preserved?.after ?? []) table.add(child.clone());
  return table;
}

/**
 * Whether the last row needs a plain row written after it. A reader takes the blank rows that end a
 * table for LibreOffice's padding and ignores their height, so a resized row with no cells would
 * reopen at the default height if it came last.
 */
function needsRowAfter(sheet: Worksheet, lastRow: number, lastColumn: number): boolean {
  const model = sheet.findRow(lastRow);
  if (!model || model.height === Defaults.rowHeight || lastRow + 1 >= SheetDimensions.maxRows) return false;
  return sheet.cellsIn(CellRange.fromBounds(lastRow, 0, lastRow, lastColumn)).length === 0;
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
      const layout = columns ? null : rowModelLayout(sheet.findRow(row));
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
 * One row: its layout, then its cells left to right. Columns are visited in segments — a stored
 * cell, or a run with the same validation and merge state — so a blank row costs a few elements
 * however wide the sheet. Neighbouring blank cells that came out identical are folded into one
 * repeated cell.
 */
function writeRow(w: SheetWriter, row: number, cellColumns: readonly number[]): XmlElement {
  const { sheet, registry } = w;
  const element = new XmlElement(Table('table-row'));
  writeRowLayout(element, sheet.findRow(row), registry);

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

  if (cell) writeCellStyle(element, cell.styleId, w.registry);

  if (validationName !== null) element.setAttr(Table(N.contentValidationName), validationName);

  writeMergeSpans(element, merge, row, column);

  if (cell) {
    writeCellChildren(element, cell.unmodeled, () => {
      const checkbox = covered ? undefined : w.checkboxes.get(cellKey(row, column));
      if (checkbox !== undefined) element.add(checkboxControl(w.sheet, row, column, checkbox));
      // A cell without a style of its own shows its row's or column's number format, so its value
      // is typed by that format too: a date under a date column is an office:date-value, as
      // LibreOffice writes it, not a float whose text is the serial.
      writeCellContent(element, cell, w.registry, w.sheet.effectiveStyleId(row, column));
    });
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
