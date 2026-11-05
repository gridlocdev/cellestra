import type { CellRange } from '../model/cellRange.ts';

/** One sort key: a zero-based column index within the sorted range's sheet, plus direction. */
export interface SortLevel {
  readonly column: number;
  readonly ascending: boolean;
}

export interface SortSpec {
  readonly levels: readonly SortLevel[];
  readonly hasHeader: boolean;
}

/**
 * One column's AutoFilter criteria: an allowed-values checklist (invariant strings) and/or a
 * condition string (`>5`, `*text*`) evaluated like SUMIF criteria.
 */
export interface ColumnFilter {
  readonly allowedValues?: ReadonlySet<string> | null;
  readonly condition?: string | null;
}

/** An AutoFilter over a range: per-column criteria keyed by absolute column index. */
export interface FilterSpec {
  readonly columns: ReadonlyMap<number, ColumnFilter>;
  readonly hasHeader: boolean;
}

/**
 * A named, auto-expanding table (PRD §8.10), persisted as ODF `table:database-range` — never an
 * Excel ListObject, so formulas use plain A1/named refs, never `Table[Column]`.
 */
export interface Table {
  readonly name: string;
  readonly range: CellRange;
  readonly hasHeader: boolean;
  readonly filter?: FilterSpec | null;
  readonly sort?: SortSpec | null;
  /** First column reads as row labels (bolded cells + theme band). Round-trips as a Cellestra attribute. */
  readonly hasHeaderColumn?: boolean;
  /** Stripe alternate data rows — pure chrome, opt-in. */
  readonly hasBandedRows?: boolean;
}

export type ValidationKind = 'any' | 'list' | 'wholeNumber' | 'decimal' | 'date' | 'textLength';
export type ValidationOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterOrEqual'
  | 'lessOrEqual';
/** How invalid input is handled; stop is the default (PRD §8.13). */
export type ValidationResponse = 'stop' | 'warning' | 'information';

/** A rule bound to a range (PRD §8.13) ↔ ODF `table:content-validation`. Checked on entry only. */
export interface DataValidation {
  readonly kind: ValidationKind;
  readonly target: CellRange;
  readonly operator: ValidationOperator;
  readonly operand1: number;
  readonly operand2: number;
  readonly listValues?: readonly string[] | null;
  readonly listSource?: CellRange | null;
  readonly inputMessage?: string | null;
  readonly errorMessage?: string | null;
  readonly response: ValidationResponse;
}

/** A validation with defaults filled in. */
export function dataValidation(
  kind: ValidationKind,
  target: CellRange,
  rest: Partial<Omit<DataValidation, 'kind' | 'target'>> = {},
): DataValidation {
  return { operator: 'between', operand1: 0, operand2: 0, response: 'stop', ...rest, kind, target };
}

export function table(
  name: string,
  range: CellRange,
  rest: Partial<Omit<Table, 'name' | 'range'>> = {},
): Table {
  return { hasHeader: true, ...rest, name, range };
}
