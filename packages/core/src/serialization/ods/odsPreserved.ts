import type { DataValidation } from '../../data/types.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import type { XmlAttribute, XmlElement } from '../../xml/index.ts';

/**
 * Content read from a `.ods` file that the model has no field for, kept verbatim so a save writes it
 * back (Golden Rule #5). It hangs off the model objects it belongs to through weak maps, so it lives
 * exactly as long as they do and the model's shape is unchanged. A cell's unmodeled children are
 * the exception: they live on the cell itself (`Cell.unmodeled`), because editing commands replace
 * cells with clones and the content has to move with them.
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

/** Workbook-level content: global named expressions and the automatic styles preserved content uses. */
export interface PreservedWorkbook {
  readonly namedExpressions: readonly XmlElement[];
  readonly automaticStyles: readonly XmlElement[];
  /** content.xml's `office:font-face-decls`, which the automatic styles written back may name. */
  readonly fontFaces?: XmlElement | null;
}

/** A validation whose condition the model cannot express, kept as the original element. */
export interface PreservedValidation {
  readonly element: XmlElement;
}

const tables = new WeakMap<Worksheet, PreservedTable>();
const workbooks = new WeakMap<Workbook, PreservedWorkbook>();
const validations = new WeakMap<DataValidation, PreservedValidation>();
const validationSources = new WeakMap<DataValidation, XmlElement>();
const databaseRangeSources = new WeakMap<object, XmlElement>();

export const OdsPreserved = {
  setTable(sheet: Worksheet, content: PreservedTable): void {
    tables.set(sheet, content);
  },
  table(sheet: Worksheet): PreservedTable | undefined {
    return tables.get(sheet);
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

  /**
   * The `table:content-validation` a modeled rule was read from: the writer takes from it what the
   * model has no field for (allow-empty-cell, display-list, message titles and display flags). An
   * edited rule is a new object, so it starts from the defaults again.
   */
  setValidationSource(validation: DataValidation, element: XmlElement): void {
    validationSources.set(validation, element);
  },
  validationSource(validation: DataValidation): XmlElement | undefined {
    return validationSources.get(validation);
  },

  /**
   * The `table:database-range` a table (or a sheet's anonymous range, keyed by its `activeFilter`
   * range object) was read from, for what the model has no field for: whether it shows filter
   * buttons, its data source, subtotal rules, sort options. A table or range replaced by an edit is
   * a new object, so it starts from the defaults again.
   */
  setDatabaseRangeSource(owner: object, element: XmlElement): void {
    databaseRangeSources.set(owner, element);
  },
  databaseRangeSource(owner: object): XmlElement | undefined {
    return databaseRangeSources.get(owner);
  },
} as const;
