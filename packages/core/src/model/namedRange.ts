import { CellAddress } from './cellAddress.ts';
import type { CellRange } from './cellRange.ts';
import type { Workbook } from './workbook.ts';

/**
 * A name in the workbook (PRD §8). It holds either a range (`Prices` = `Data!C3:C9`) or an
 * expression (`TaxRate` = `0.2`: formula source without `=`, in the notation cell formulas use).
 * `scope` null = workbook-global, else the owning sheet's name. An unqualified reference in an
 * expression means the sheet of the formula that uses the name.
 */
export type NamedRange = NamedTarget | NamedExpression;

interface NameBase {
  readonly name: string;
  readonly scope?: string | null;
}

/** A name for cells. */
export interface NamedTarget extends NameBase {
  readonly target: CellRange;
  readonly expression?: undefined;
}

/** A name for a value or formula, such as `TaxRate` = `0.2`. */
export interface NamedExpression extends NameBase {
  readonly expression: string;
  readonly target?: undefined;
}

/** What a name stands for, as the Name Manager and reports show it: `A1:B2` or `=0.2`. */
export function namedRangeDefinition(named: NamedRange): string {
  return named.target ? named.target.toString() : `=${named.expression}`;
}

const NAME = /^[A-Za-z_\\][A-Za-z0-9_.]*$/;

/**
 * Why `name` can't be added to the workbook, or null when it can: a name starts with a letter,
 * underscore or backslash, can't look like a cell reference, and is unique in its scope.
 */
export function namedRangeError(
  workbook: Workbook,
  name: string,
  scope: string | null = null,
): string | null {
  if (!NAME.test(name) || CellAddress.tryParse(name))
    return 'Names start with a letter or underscore and can’t look like a cell reference.';
  if (workbook.findNamedRange(name, scope)) return `“${name}” already exists.`;
  return null;
}
