import {
  table as makeTable,
  type ColumnFilter,
  type FilterSpec,
  type SortLevel,
  type SortSpec,
  type Table as DataTable,
} from '../../data/types.ts';
import { OpenFormulaTranslator } from '../../formulas/openFormulaTranslator.ts';
import { CellAddress } from '../../model/cellAddress.ts';
import { CellRange } from '../../model/cellRange.ts';
import type { Workbook } from '../../model/workbook.ts';
import { XmlElement } from '../../xml/index.ts';
import { odfRangeAddress } from './odsContentShared.ts';
import { OdsPreserved } from './odsPreserved.ts';
import { CellestraNs, OdsXmlNames as N, Table } from './odsXmlNames.ts';

/**
 * LibreOffice names a sheet's anonymous AutoFilter range `__Anonymous_Sheet_DB__<sheet index>`, one
 * per sheet; a single shared name made every sheet after the first collide with the first on open.
 * Older Cellestra files wrote the bare `__Anonymous_Sheet__`, still read as the same thing.
 */
const ANONYMOUS_DATABASE_RANGE_PREFIX = '__Anonymous_Sheet_DB__';
const LEGACY_ANONYMOUS_DATABASE_RANGE_NAME = '__Anonymous_Sheet__';

function isAnonymousDatabaseRange(name: string): boolean {
  return name.startsWith(ANONYMOUS_DATABASE_RANGE_PREFIX) || name === LEGACY_ANONYMOUS_DATABASE_RANGE_NAME;
}

const CONDITION_OPERATORS: ReadonlyArray<readonly [prefix: string, operator: string]> = [
  ['>=', '>='],
  ['<=', '<='],
  ['<>', '!='],
  ['>', '>'],
  ['<', '<'],
  ['=', '='],
];

// ---- table:database-range (filter + sort, PRD §8.10–§8.12) ----------------

/**
 * Emit one `table:database-range` per worksheet table and per sheet-level AutoFilter, carrying the
 * filter region, the active filter conditions, and the sort order so they round-trip. Filtering
 * sets row visibility (written by the row writer as `table:visibility`).
 */
export function addDatabaseRanges(spreadsheet: XmlElement, workbook: Workbook): void {
  const ranges: XmlElement[] = [];
  for (const [sheetIndex, sheet] of workbook.worksheets.entries()) {
    for (const t of sheet.tables) {
      const element = writeDatabaseRange(
        sheet.name,
        t.name,
        t.range,
        t.hasHeader,
        t.filter ?? null,
        t.sort ?? null,
        OdsPreserved.databaseRangeSource(t) ?? null,
      );
      // Facts ODF has no attribute for ride along as cellestra:* foreign attributes, only when on.
      if (t.hasHeaderColumn) element.setAttr(CellestraNs(N.headerColumn), 'true');
      if (t.hasBandedRows) element.setAttr(CellestraNs(N.bandedRows), 'true');
      ranges.push(element);
    }

    if (sheet.activeFilter) {
      const hasHeader = sheet.activeFilterSpec?.hasHeader ?? sheet.activeSort?.hasHeader ?? true;
      ranges.push(
        writeDatabaseRange(
          sheet.name,
          `${ANONYMOUS_DATABASE_RANGE_PREFIX}${sheetIndex}`,
          sheet.activeFilter,
          hasHeader,
          sheet.activeFilterSpec,
          sheet.activeSort,
          OdsPreserved.databaseRangeSource(sheet.activeFilter) ?? null,
        ),
      );
    }
  }

  if (ranges.length === 0) return;
  spreadsheet.add(new XmlElement(Table(N.databaseRanges), ...ranges));
}

function writeDatabaseRange(
  sheetName: string,
  name: string,
  range: CellRange,
  hasHeader: boolean,
  filter: FilterSpec | null,
  sort: SortSpec | null,
  source: XmlElement | null,
): XmlElement {
  const element = new XmlElement(Table(N.databaseRange))
    .setAttr(Table(N.targetRangeAddress), odfRangeAddress(range.withSheet(sheetName), sheetName))
    .setAttr(Table(N.displayFilterButtons), 'true')
    .setAttr(Table(N.containsHeader), hasHeader ? 'true' : 'false')
    .setAttr(Table('name'), name);

  if (source) {
    // A range read from a file keeps what the model does not own. A range LibreOffice made for a
    // sort has no filter buttons (the attribute's default is false), and must not gain them.
    element.setAttr(Table(N.displayFilterButtons), source.attr(Table(N.displayFilterButtons)) ?? 'false');
    for (const a of source.attributes) {
      if (a.name.uri === CellestraNs.uri) continue;
      if (a.name.uri === Table.uri && OWNED_ATTRIBUTES.has(a.name.local)) continue;
      element.setAttr(a.name, a.value);
    }
  }

  // ODF order: a data source, then the filter, the sort, and subtotal rules.
  const kept = source?.elements().filter((e) => !e.is(Table(N.sort)) && !e.is(Table(N.filter))) ?? [];
  const isSource = (e: XmlElement): boolean => e.uri === Table.uri && e.local.startsWith('database-source');
  for (const e of kept) if (isSource(e)) element.add(e.clone());
  if (filter && filter.columns.size > 0) {
    const written = writeFilter(filter, range.left);
    copyUnowned(source?.element(Table(N.filter)) ?? null, written);
    element.add(written);
  }
  if (sort && sort.levels.length > 0) {
    const written = writeSort(sort, range.left);
    const sourceSort = source?.element(Table(N.sort)) ?? null;
    copyUnowned(sourceSort, written);
    // A key's data type (text, number, a user list) stays with its field.
    for (const by of written.elements(Table(N.sortBy))) {
      const field = by.attr(Table(N.fieldNumber));
      const from = sourceSort?.elements(Table(N.sortBy)).find((b) => b.attr(Table(N.fieldNumber)) === field);
      copyUnowned(from ?? null, by);
    }
    element.add(written);
  }
  for (const e of kept) if (!isSource(e)) element.add(e.clone());
  return element;
}

/** The database-range attributes the model owns (written from it, never copied from the file). */
const OWNED_ATTRIBUTES: ReadonlySet<string> = new Set([
  N.targetRangeAddress,
  N.containsHeader,
  N.displayFilterButtons,
  'name',
]);

/** Copy `from`'s attributes that `to` does not already carry. */
function copyUnowned(from: XmlElement | null, to: XmlElement): void {
  for (const a of from?.attributes ?? []) if (to.attr(a.name) === null) to.setAttr(a.name, a.value);
}

function writeSort(sort: SortSpec, rangeLeft: number): XmlElement {
  const element = new XmlElement(Table(N.sort));
  for (const level of sort.levels) {
    element.add(
      new XmlElement(Table(N.sortBy))
        .setAttr(Table(N.fieldNumber), level.column - rangeLeft)
        .setAttr(Table(N.order), level.ascending ? 'ascending' : 'descending'),
    );
  }
  return element;
}

function writeFilter(filter: FilterSpec, rangeLeft: number): XmlElement {
  // AND across columns; within a column, a value checklist is OR'd '=' conditions.
  const and = new XmlElement(Table(N.filterAnd));
  const columns = [...filter.columns].sort((a, b) => a[0] - b[0]);
  for (const [column, columnFilter] of columns) {
    const field = column - rangeLeft;

    const allowed = columnFilter.allowedValues;
    if (allowed && allowed.size > 0) {
      if (allowed.size === 1) {
        and.add(condition(field, '=', [...allowed][0]!));
      } else {
        const or = new XmlElement(Table(N.filterOr));
        for (const value of allowed) or.add(condition(field, '=', value));
        and.add(or);
      }
    }

    if (columnFilter.condition) {
      const [op, value] = splitCondition(columnFilter.condition);
      and.add(condition(field, op, value, isNumericCondition(op, value)));
    }
  }
  return new XmlElement(Table(N.filter), and);
}

function condition(field: number, op: string, value: string, numeric = false): XmlElement {
  const element = new XmlElement(Table(N.filterCondition))
    .setAttr(Table(N.fieldNumber), field)
    .setAttr(Table(N.operator), op)
    .setAttr(Table('value'), value);
  // Without a data type a condition compares as text, so `>5` matched "6" but not 10 ("10" < "5").
  if (numeric) element.setAttr(Table('data-type'), 'number');
  return element;
}

const NUMERIC_OPERATORS: ReadonlySet<string> = new Set(['<', '>', '<=', '>=', '=', '!=']);

function isNumericCondition(op: string, value: string): boolean {
  return NUMERIC_OPERATORS.has(op) && /^\s*[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?\s*$/.test(value);
}

/** Split a SUMIF-style criteria string (`>5`, `<=10`, `<>x`, plain `text`) into an ODF operator + value. */
function splitCondition(text: string): [string, string] {
  for (const [prefix, op] of CONDITION_OPERATORS) {
    if (text.startsWith(prefix)) return [op, text.slice(prefix.length)];
  }
  return ['=', text];
}

function joinCondition(op: string, value: string): string {
  if (op === '=') return value;
  if (op === '!=') return '<>' + value;
  return op + value;
}

// ---- table:database-range (read side) -------------------------------------

export function readDatabaseRanges(container: XmlElement, workbook: Workbook): void {
  for (const element of container.elements(Table(N.databaseRange))) readDatabaseRange(element, workbook);
}

function readDatabaseRange(element: XmlElement, workbook: Workbook): void {
  const address = element.attr(Table(N.targetRangeAddress));
  if (address === null) return;

  const internalRef = OpenFormulaTranslator.fromOpenFormula('of:=[' + address.replace(/^\$+/, '') + ']');
  const range = CellRange.tryParse(internalRef);
  if (!range || range.sheet === null) return;

  const sheet = workbook.findSheet(range.sheet);
  if (!sheet) return;

  const hasHeader = element.attr(Table(N.containsHeader)) !== 'false';

  // Normalize to a plain relative, sheet-less range so it compares equal to one built in code.
  const localRange = new CellRange(
    new CellAddress(range.left, range.top),
    new CellAddress(range.right, range.bottom),
  );

  const sort = readSort(element, localRange.left, hasHeader);
  const filter = readFilter(element, localRange.left, hasHeader);

  const name = element.attr(Table('name'));
  if (name !== null && !isAnonymousDatabaseRange(name)) {
    const hasHeaderColumn = element.attr(CellestraNs(N.headerColumn)) === 'true';
    const hasBandedRows = element.attr(CellestraNs(N.bandedRows)) === 'true';

    // A named database range is a Cellestra table; merge filter/sort onto it (or add it).
    const existing = sheet.tables.find((t) => t.name === name);
    let added: DataTable | null = null;
    if (existing) {
      added = { ...existing, filter, sort, hasHeaderColumn, hasBandedRows };
      sheet.replaceTable(existing, added);
    } else if (!sheet.tables.some((t) => t.range.intersects(localRange))) {
      added = makeTable(name, localRange, { hasHeader, filter, sort, hasHeaderColumn, hasBandedRows });
      sheet.addTable(added);
    }
    if (added) OdsPreserved.setDatabaseRangeSource(added, element.clone());
  } else {
    // An anonymous database range is the sheet-level AutoFilter region.
    sheet.activeFilter = localRange;
    sheet.activeFilterSpec = filter;
    sheet.activeSort = sort;
    OdsPreserved.setDatabaseRangeSource(localRange, element.clone());
  }

  // The filter isn't re-run here. Filtered rows still come back hidden because the row writer
  // persisted their `table:visibility`, so only a file whose visibility disagrees with its filter
  // would show different rows.
}

function readSort(databaseRange: XmlElement, rangeLeft: number, hasHeader: boolean): SortSpec | null {
  const sortElement = databaseRange.element(Table(N.sort));
  if (!sortElement) return null;

  const levels: SortLevel[] = [];
  for (const by of sortElement.elements(Table(N.sortBy))) {
    const field = parseInt32(by.attr(Table(N.fieldNumber)));
    if (field !== null) {
      levels.push({ column: rangeLeft + field, ascending: by.attr(Table(N.order)) !== 'descending' });
    }
  }
  return levels.length > 0 ? { levels, hasHeader } : null;
}

function readFilter(databaseRange: XmlElement, rangeLeft: number, hasHeader: boolean): FilterSpec | null {
  const filterElement = databaseRange.element(Table(N.filter));
  const top = filterElement?.element(Table(N.filterAnd)) ?? filterElement?.element(Table(N.filterOr)) ?? null;
  if (!top) return null;

  // Per field, gather '=' values into a checklist and the first relational op into a condition.
  const allowed = new Map<number, Set<string>>();
  const conditions = new Map<number, string>();

  const absorb = (conditionElement: XmlElement): void => {
    const field = parseInt32(conditionElement.attr(Table(N.fieldNumber)));
    if (field === null) return;
    const op = conditionElement.attr(Table(N.operator)) ?? '=';
    const value = conditionElement.attr(Table('value')) ?? '';
    const column = rangeLeft + field;

    if (op === '=') {
      let set = allowed.get(column);
      if (!set) {
        set = new Set();
        allowed.set(column, set);
      }
      set.add(value);
    } else if (!conditions.has(column)) {
      conditions.set(column, joinCondition(op, value));
    }
  };

  for (const node of top.elements()) {
    if (node.is(Table(N.filterCondition))) absorb(node);
    else if (node.is(Table(N.filterOr)) || node.is(Table(N.filterAnd))) {
      for (const nested of node.elements(Table(N.filterCondition))) absorb(nested);
    }
  }

  const columns = new Map<number, ColumnFilter>();
  for (const column of new Set([...allowed.keys(), ...conditions.keys()])) {
    columns.set(column, {
      allowedValues: allowed.get(column) ?? null,
      condition: conditions.get(column) ?? null,
    });
  }
  return columns.size > 0 ? { columns, hasHeader } : null;
}

/** `int.TryParse` (invariant): optional sign and digits, surrounding whitespace allowed. */
function parseInt32(text: string | null): number | null {
  if (text === null) return null;
  const t = text.trim();
  if (!/^[+-]?\d+$/.test(t)) return null;
  const n = Number(t);
  return n >= -2147483648 && n <= 2147483647 ? n : null;
}
