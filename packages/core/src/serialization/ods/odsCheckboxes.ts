import { OpenFormulaTranslator } from '../../formulas/openFormulaTranslator.ts';
import { CellAddress } from '../../model/cellAddress.ts';
import { cellKey, SheetDimensions } from '../../model/sheetDimensions.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { xmlNamespace, XmlElement, XML_NS_URI, type XmlName } from '../../xml/index.ts';
import { formatUpTo3, odfSheetName, pxToCm } from './odsContentShared.ts';
import { Draw, Office, Svg, Table } from './odsXmlNames.ts';

/**
 * Cell checkboxes as standard ODF form controls, so LibreOffice shows (and toggles) the same
 * checkbox Cellestra does. Each one is a `form:checkbox` in the sheet's `office:forms`, linked to
 * its cell with `form:linked-cell` (clicking it writes TRUE/FALSE there), drawn by a `draw:control`
 * anchored in that cell (ODF 1.3 §13.5.4, §10.3.13). The model keeps only the cell's `checkbox`
 * style flag; the controls are regenerated from it on every save, and on read a plain linked,
 * unlabelled checkbox becomes that flag again. Anything else in the forms is preserved verbatim.
 */

export const Form = xmlNamespace('urn:oasis:names:tc:opendocument:xmlns:form:1.0', 'form');
/**
 * The namespace of the `ooo:` prefix in `form:control-implementation` values. The value is a
 * namespaced name, so the prefix must be declared or LibreOffice drops the control.
 */
export const Ooo = xmlNamespace('http://openoffice.org/2004/office', 'ooo');
const XmlId: XmlName = { uri: XML_NS_URI, prefix: 'xml', local: 'id' };

const CHECKBOX_IMPLEMENTATION = 'ooo:com.sun.star.form.component.CheckBox';
const FORM_IMPLEMENTATION = 'ooo:com.sun.star.form.component.Form';
const FORM_NAME = 'Cellestra Checkboxes';

/** The checkboxes a file's tables link to cells: their control ids, and the cells to flag. */
export interface ReadCheckboxes {
  /** Control ids taken over by the model; their `form:checkbox` and `draw:control` are dropped. */
  readonly ids: Set<string>;
  readonly cells: Array<{ readonly sheet: string; readonly row: number; readonly column: number }>;
}

function controlId(e: XmlElement): string | null {
  return e.attr(XmlId) ?? e.attr(Form('id'));
}

function parseLinkedCell(text: string | null, fallbackSheet: string): CellAddress | null {
  if (!text) return null;
  const internal = OpenFormulaTranslator.fromOpenFormula('of:=[' + text.replace(/^\$+/, '') + ']');
  const address = CellAddress.tryParse(internal.replace(/^=/, ''));
  if (!address) return null;
  return address.sheet === null ? address.withSheet(fallbackSheet) : address;
}

/** Whether a `form:checkbox` is one the model can hold: linked to one cell, with no label of its own. */
function modelable(e: XmlElement, sheetName: string): CellAddress | null {
  if (!e.is(Form('checkbox')) || controlId(e) === null) return null;
  const label = e.attr(Form('label'));
  if (label !== null && label !== '') return null;
  return parseLinkedCell(e.attr(Form('linked-cell')), sheetName);
}

/**
 * Collect a table's modelable checkboxes into `into`. Run before the table's cells are read.
 *
 * Only a checkbox whose linked cell is in this table and holds TRUE or FALSE is taken over: that
 * is the only cell the model draws (and saves) a checkbox for. Any other — linked to a blank, a
 * number, a covered cell or another sheet — stays in the file exactly as it was.
 */
export function collectCheckboxes(table: XmlElement, sheetName: string, into: ReadCheckboxes): void {
  const forms = table.element(Office('forms'));
  if (!forms) return;
  const candidates: Array<{ control: XmlElement; address: CellAddress }> = [];
  for (const form of forms.elements(Form('form'))) {
    for (const control of form.elements(Form('checkbox'))) {
      const address = modelable(control, sheetName);
      if (!address || (address.sheet ?? sheetName).toLowerCase() !== sheetName.toLowerCase()) continue;
      candidates.push({ control, address });
    }
  }
  if (candidates.length === 0) return;

  const linked = cellsAt(
    table,
    new Set(candidates.map(({ address }) => cellKey(address.row, address.column))),
  );
  for (const { control, address } of candidates) {
    const cell = linked.get(cellKey(address.row, address.column));
    if (!cell?.is(Table('table-cell')) || cell.attr(Office('value-type')) !== 'boolean') continue;
    into.ids.add(controlId(control)!);
    into.cells.push({ sheet: sheetName, row: address.row, column: address.column });
  }
}

const ROW_CONTAINERS = ['table-header-rows', 'table-row-group', 'table-rows'];

/** The cell elements at the `wanted` positions (by `cellKey`), walking the grid's repeats. */
function cellsAt(table: XmlElement, wanted: ReadonlySet<number>): Map<number, XmlElement> {
  const maxRow = Math.max(...[...wanted].map((k) => Math.floor(k / SheetDimensions.maxColumns)));
  const found = new Map<number, XmlElement>();
  let row = 0;
  const stack: Array<{ e: XmlElement; i: number }> = [{ e: table, i: 0 }];
  while (stack.length > 0 && row <= maxRow) {
    const top = stack[stack.length - 1]!;
    if (top.i >= top.e.nodes.length) {
      stack.pop();
      continue;
    }
    const e = top.e.nodes[top.i++]!;
    if (!(e instanceof XmlElement) || e.uri !== Table.uri) continue;
    if (ROW_CONTAINERS.includes(e.local)) {
      stack.push({ e, i: 0 });
      continue;
    }
    if (e.local !== 'table-row') continue;
    const rows = repeatOf(e, 'number-rows-repeated');
    let column = 0;
    for (const cell of e.elements()) {
      if (!cell.is(Table('table-cell')) && !cell.is(Table('covered-table-cell'))) continue;
      const columns = repeatOf(cell, 'number-columns-repeated');
      for (const key of wanted) {
        const r = Math.floor(key / SheetDimensions.maxColumns);
        const c = key % SheetDimensions.maxColumns;
        if (r >= row && r < row + rows && c >= column && c < column + columns) found.set(key, cell);
      }
      column += columns;
    }
    row += rows;
  }
  return found;
}

function repeatOf(e: XmlElement, attribute: string): number {
  const n = Number(e.attr(Table(attribute)) ?? '1');
  return Number.isInteger(n) && n > 0 ? n : 1;
}

/** A table child with the taken-over checkboxes removed; null when nothing is left of it. */
export function withoutCheckboxes(child: XmlElement, ids: ReadonlySet<string>): XmlElement | null {
  if (ids.size === 0) return child;
  if (child.is(Office('forms'))) {
    const copy = child.clone();
    for (const form of copy.elements(Form('form'))) {
      for (const control of form.elements(Form('checkbox'))) {
        const id = controlId(control);
        if (id !== null && ids.has(id)) form.remove(control);
      }
      if (form.elements().every((e) => e.is(Form('properties')) || e.is(Office('event-listeners'))))
        copy.remove(form);
    }
    return copy.hasElements ? copy : null;
  }
  if (child.is(Table('shapes'))) {
    const copy = child.clone();
    for (const shape of copy.elements()) if (isTakenControl(shape, ids)) copy.remove(shape);
    return copy.hasElements ? copy : null;
  }
  return child;
}

/** A `draw:control` that draws one of the taken-over checkboxes. */
export function isTakenControl(e: XmlElement, ids: ReadonlySet<string>): boolean {
  if (!e.is(Draw('control'))) return false;
  const id = e.attr(Draw('control'));
  return id !== null && ids.has(id);
}

/** Flag each linked cell's style as a checkbox. Run once every table has been read. */
export function applyCheckboxes(workbook: Workbook, read: ReadCheckboxes): void {
  for (const { sheet: name, row, column } of read.cells) {
    const sheet = workbook.findSheet(name);
    if (!sheet) continue;
    const base = sheet.effectiveStyleId(row, column);
    const cell = sheet.getOrCreate(row, column);
    cell.styleId = workbook.styles.mutate(base, (s) => (s.checkbox ? s : { ...s, checkbox: true }));
  }
}

/** The cells of `sheet` that save as a checkbox, keyed by `cellKey`, each with its control id. */
export function checkboxCells(workbook: Workbook, sheet: Worksheet, sheetIndex: number): Map<number, string> {
  const keys: number[] = [];
  // Most workbooks have no checkbox at all; then there is nothing to look for cell by cell.
  if (!workbook.styles.styles.some((s) => s.checkbox)) return new Map();
  for (const [row, column, cell] of sheet.enumerateCells()) {
    // Only a TRUE/FALSE cell shows a box (as in the grid); a cleared one is a plain cell again.
    if (cell.value.kind !== 'boolean') continue;
    if (!workbook.styles.get(sheet.effectiveStyleId(row, column)).checkbox) continue;
    const merge = sheet.findMerge(row, column);
    if (merge && (merge.range.top !== row || merge.range.left !== column)) continue;
    keys.push(cellKey(row, column));
  }
  // Row by row, so the ids and the form's tab order follow the sheet.
  keys.sort((a, b) => a - b);
  return new Map(keys.map((k, i) => [k, `cellestra-checkbox-${sheetIndex + 1}-${i + 1}`]));
}

/** The `form:form` holding a sheet's checkboxes, in `checkboxCells` order. */
export function checkboxForm(sheet: Worksheet, cells: ReadonlyMap<number, string>): XmlElement {
  const form = new XmlElement(Form('form'))
    .setAttr(Form('name'), FORM_NAME)
    .setAttr(Form('control-implementation'), FORM_IMPLEMENTATION)
    .setAttr(Form('apply-filter'), 'true')
    .setAttr(Form('command-type'), 'table');
  for (const [key, id] of cells) {
    const row = Math.floor(key / SheetDimensions.maxColumns);
    const column = key % SheetDimensions.maxColumns;
    const cell = sheet.find(row, column);
    const address = new CellAddress(column, row);
    const checked = cell?.value.kind === 'boolean' && cell.value.value;
    form.add(
      new XmlElement(Form('checkbox'))
        .setAttr(Form('name'), `Checkbox ${address.columnLabel}${address.rowLabel}`)
        .setAttr(Form('control-implementation'), CHECKBOX_IMPLEMENTATION)
        .setAttr(XmlId, id)
        .setAttr(Form('id'), id)
        .setAttr(Form('linked-cell'), relativeAddress(sheet.name, row, column))
        .setAttr(Form('current-state'), checked ? 'checked' : 'unchecked')
        .setAttr(Form('image-position'), 'center'),
    );
  }
  return form;
}

/** A new `office:forms` around `form`, live (not in design mode) so the checkboxes can be clicked. */
export function formsElement(form: XmlElement): XmlElement {
  return new XmlElement(Office('forms'))
    .setAttr(Form('automatic-focus'), 'false')
    .setAttr(Form('apply-design-mode'), 'false')
    .add(form);
}

/**
 * `'My Sheet'.B2`: the relative form LibreOffice writes for these two attributes. It reads a
 * `$`-absolute `table:end-cell-address` as A1, which stretches the control from the sheet corner.
 */
function relativeAddress(sheet: string, row: number, column: number): string {
  return `${odfSheetName(sheet)}.${SheetDimensions.columnToLetters(column)}${row + 1}`;
}

const cm = (px: number) => formatUpTo3(pxToCm(px)) + 'cm';

/** The `draw:control` that draws checkbox `id` over the cell at (row, column), anchored to it. */
export function checkboxControl(sheet: Worksheet, row: number, column: number, id: string): XmlElement {
  // A merged checkbox spans the merge, so the control ends in the merge's last cell.
  const merge = sheet.findMerge(row, column);
  const bottom = merge?.range.bottom ?? row;
  const right = merge?.range.right ?? column;
  let width = 0;
  for (let c = column; c <= right; c++) width += sheet.isColumnHidden(c) ? 0 : sheet.getColumnWidth(c);
  let height = 0;
  for (let r = row; r <= bottom; r++) height += sheet.isRowHidden(r) ? 0 : sheet.getRowHeight(r);
  // One pixel in from each edge, so the control sits inside the cell's grid lines. LibreOffice
  // measures `svg:x`/`svg:y` from the anchor cell and `table:end-x`/`end-y` from the end cell.
  const inset = Math.min(1, width / 4, height / 4);
  return new XmlElement(Draw('control'))
    .setAttr(Table('end-cell-address'), relativeAddress(sheet.name, bottom, right))
    .setAttr(Table('end-x'), cm(sheet.getColumnWidth(right) - inset))
    .setAttr(Table('end-y'), cm(sheet.getRowHeight(bottom) - inset))
    .setAttr(Draw('z-index'), '0')
    .setAttr(Svg('width'), cm(Math.max(0, width - 2 * inset)))
    .setAttr(Svg('height'), cm(Math.max(0, height - 2 * inset)))
    .setAttr(Svg('x'), cm(inset))
    .setAttr(Svg('y'), cm(inset))
    .setAttr(Draw('control'), id);
}
