import type { DataValidation } from '../../data/types.ts';
import type { Cell } from '../../model/cell.ts';
import { cellKey } from '../../model/sheetDimensions.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import type { XmlAttribute, XmlElement } from '../../xml/index.ts';

/**
 * Content read from a `.ods` file that the model has no field for, kept verbatim so a save writes it
 * back (Golden Rule #5). It hangs off the model objects it belongs to through weak maps, so it lives
 * exactly as long as they do and the model's shape is unchanged.
 *
 * Every element stored here is a detached clone; the writer clones it again before adding it to a
 * new tree, so one read can be saved any number of times.
 */

/** A table's unmodeled attributes and children, split by where they sat relative to its grid. */
export interface PreservedTable {
  /** Every `table:table` attribute except `table:name`, which the writer owns. */
  readonly attributes: readonly XmlAttribute[];
  /** Children before the first column or row (`table:shapes`, `office:forms`, …). */
  readonly before: readonly XmlElement[];
  /** Children after the grid (`calcext:conditional-formats`, …). */
  readonly after: readonly XmlElement[];
  /** Sheet-local `table:named-expressions` entries the model does not hold (named expressions). */
  readonly namedExpressions: readonly XmlElement[];
}

/** A cell's unmodeled children (`office:annotation`, `draw:frame`, …), split around its text. */
export interface PreservedCell {
  readonly before: readonly XmlElement[];
  readonly after: readonly XmlElement[];
}

/** Workbook-level content: global named expressions and the automatic styles preserved content uses. */
export interface PreservedWorkbook {
  readonly namedExpressions: readonly XmlElement[];
  readonly automaticStyles: readonly XmlElement[];
}

/** A validation whose condition the model cannot express, kept as the original element. */
export interface PreservedValidation {
  readonly element: XmlElement;
}

interface PositionedCell {
  readonly fingerprint: string;
  readonly content: PreservedCell;
}

const tables = new WeakMap<Worksheet, PreservedTable>();
const cells = new WeakMap<Cell, PreservedCell>();
/**
 * Editing commands (sort, insert/delete, undo) replace cells with clones, which the weak map does
 * not follow. As a fallback a cell's content is also kept by position, and written back to whatever
 * cell sits there if it still holds the same value and formula.
 */
const cellsByPosition = new WeakMap<Worksheet, Map<number, PositionedCell>>();
const workbooks = new WeakMap<Workbook, PreservedWorkbook>();
const validations = new WeakMap<DataValidation, PreservedValidation>();

function fingerprint(cell: Cell): string {
  const v = cell.value;
  const value =
    v.kind === 'number'
      ? `n${v.value}`
      : v.kind === 'text'
        ? `t${v.value}`
        : v.kind === 'boolean'
          ? `b${v.value}`
          : v.kind === 'error'
            ? `e${v.error.token}`
            : '-';
  return `${value}\u0000${cell.formula ?? ''}`;
}

export const OdsPreserved = {
  setTable(sheet: Worksheet, content: PreservedTable): void {
    tables.set(sheet, content);
  },
  table(sheet: Worksheet): PreservedTable | undefined {
    return tables.get(sheet);
  },

  setCell(sheet: Worksheet, row: number, column: number, cell: Cell, content: PreservedCell): void {
    cells.set(cell, content);
    let byPosition = cellsByPosition.get(sheet);
    if (!byPosition) cellsByPosition.set(sheet, (byPosition = new Map()));
    byPosition.set(cellKey(row, column), { fingerprint: fingerprint(cell), content });
  },
  cell(sheet: Worksheet, row: number, column: number, cell: Cell): PreservedCell | undefined {
    const direct = cells.get(cell);
    if (direct) return direct;
    const positioned = cellsByPosition.get(sheet)?.get(cellKey(row, column));
    return positioned && positioned.fingerprint === fingerprint(cell) ? positioned.content : undefined;
  },

  setWorkbook(workbook: Workbook, content: PreservedWorkbook): void {
    workbooks.set(workbook, content);
  },
  workbook(workbook: Workbook): PreservedWorkbook | undefined {
    return workbooks.get(workbook);
  },

  setValidation(validation: DataValidation, content: PreservedValidation): void {
    validations.set(validation, content);
  },
  validation(validation: DataValidation): PreservedValidation | undefined {
    return validations.get(validation);
  },
} as const;
