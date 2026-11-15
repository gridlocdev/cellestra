import { Defaults } from '../../defaults.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { StylePool } from '../../styling/stylePool.ts';
import { XmlElement } from '../../xml/index.ts';
import { isHidden, readRowVisibility, writeVisibility } from './odsHiddenLines.ts';
import type { LineContainers } from './odsLineContainers.ts';
import { isSpanned } from './odsMerges.ts';
import type { OdsStyleIndex } from './odsStyleIndex.ts';
import type { StyleRegistry } from './odsStyleRegistry.ts';
import { Table } from './odsXmlNames.ts';

/**
 * Row and column layout: widths and heights (through their automatic styles), page breaks, and a
 * whole line's cell format (`table:default-cell-style-name`). Visibility is `odsHiddenLines`.
 */

// ── Reading ─────────────────────────────────────────────────────────────────

/** A line's `table:default-cell-style-name`, as a style id; null when it has none or the default. */
export function defaultStyleIdOf(element: XmlElement, styles: OdsStyleIndex): number | null {
  const name = element.attr(Table('default-cell-style-name'));
  if (name === null) return null;
  const id = styles.cellStyleId(name);
  return id === StylePool.defaultId ? null : id;
}

/** Apply a `table:table-column` element's layout to columns `column … column + repeat - 1`. */
export function readColumnLayout(
  sheet: Worksheet,
  element: XmlElement,
  styles: OdsStyleIndex,
  column: number,
  repeat: number,
): void {
  const styleName = element.attr(Table('style-name'));
  const hidden = isHidden(element);
  const width = styleName === null ? null : styles.columnWidth(styleName);
  const pageBreak = styleName !== null && styles.pageBreak(styleName);
  // The column's default cell style: "this whole column is formatted like so".
  const defaultStyleId = defaultStyleIdOf(element, styles);
  if (width === null && !hidden && !pageBreak && defaultStyleId === null) return;
  for (let c = column; c < column + repeat; c++) {
    const model = sheet.getOrCreateColumn(c);
    if (width !== null) model.width = width;
    if (defaultStyleId !== null) model.styleId = defaultStyleId;
    model.hidden = hidden;
    model.pageBreak = pageBreak;
  }
}

/** Apply a `table:table-row` element's layout to rows `from … to - 1`. */
export function readRowLayout(
  sheet: Worksheet,
  element: XmlElement,
  styles: OdsStyleIndex,
  from: number,
  to: number,
): void {
  const styleName = element.attr(Table('style-name'));
  const visibility = readRowVisibility(element);
  const hidden = visibility.filtered || visibility.manuallyHidden;
  const height = styleName === null ? null : styles.rowHeight(styleName);
  const pageBreak = styleName !== null && styles.pageBreak(styleName);
  const defaultStyleId = defaultStyleIdOf(element, styles);
  if (height === null && !hidden && !pageBreak && defaultStyleId === null) return;
  for (let r = from; r < to; r++) {
    const model = sheet.getOrCreateRow(r);
    if (height !== null) model.height = height;
    model.pageBreak = pageBreak;
    if (defaultStyleId !== null) model.styleId = defaultStyleId;
    model.filtered = visibility.filtered;
    model.manuallyHidden = visibility.manuallyHidden;
  }
}

/**
 * Whether an inert row element (no cell content) carries any layout. With `skipHeight`, a height
 * alone doesn't count: the padding at the end of a sheet only ever carries the default height.
 */
export function rowCarriesLayout(element: XmlElement, styles: OdsStyleIndex, skipHeight: boolean): boolean {
  const styleName = element.attr(Table('style-name'));
  const height = styleName === null ? null : styles.rowHeight(styleName);
  const pageBreak = styleName !== null && styles.pageBreak(styleName);
  return (
    isHidden(element) ||
    pageBreak ||
    defaultStyleIdOf(element, styles) !== null ||
    (height !== null && !skipHeight)
  );
}

/** What makes padding rows alike: their style, visibility and default cell style. */
export function rowElementLayout(e: XmlElement): string {
  return `${e.attr(Table('style-name'))}|${e.attr(Table('visibility'))}|${e.attr(Table('default-cell-style-name'))}`;
}

// ── Writing ─────────────────────────────────────────────────────────────────

/** Write columns `0 … lastColumn`, coalescing identical neighbours inside one group. */
export function writeColumns(
  containers: LineContainers,
  sheet: Worksheet,
  lastColumn: number,
  registry: StyleRegistry,
): void {
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
    writeVisibility(element, { hidden: signature.hidden });
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

/** The row model fields the writer reads. */
interface RowModel {
  readonly height: number;
  readonly manuallyHidden: boolean;
  readonly filtered: boolean;
  readonly pageBreak: boolean;
  readonly styleId: number;
}

/** Put a row model's layout on a `table:table-row` element: height, page break, visibility, format. */
export function writeRowLayout(
  element: XmlElement,
  model: RowModel | undefined,
  registry: StyleRegistry,
): void {
  if (!model) return;
  if (model.height !== Defaults.rowHeight || model.pageBreak) {
    element.setAttr(Table('style-name'), registry.rowStyle(model.height, model.pageBreak));
  }
  writeVisibility(element, { filtered: model.filtered, hidden: model.manuallyHidden });
  // A whole formatted row travels the same way a formatted column does — one attribute.
  if (model.styleId !== StylePool.defaultId) {
    element.setAttr(Table('default-cell-style-name'), registry.cellStyle(model.styleId));
  }
}

/** What decides how an empty row is written: its height, visibility and default style. */
export function rowModelLayout(model: RowModel | undefined): string {
  return model
    ? `${model.height}|${model.manuallyHidden}|${model.filtered}|${model.pageBreak}|${model.styleId}`
    : '';
}

/**
 * Appends rows to a table (or the group they are in), folding each into the previous one when they
 * are interchangeable and in the same group.
 */
export class RowAppender {
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

/** An element's attributes as text, but for the `table:` one named `skip`: equal when interchangeable. */
export function attributeSignature(element: XmlElement, skip: string): string {
  let s = '';
  for (const a of element.attributes) {
    if (a.name.local === skip && a.name.uri === Table.uri) continue;
    s += ` ${a.name.uri}:${a.name.local}=${a.value}`;
  }
  return s;
}
