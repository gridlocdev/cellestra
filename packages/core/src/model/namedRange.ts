import { CellAddress } from './cellAddress.ts';
import type { CellRange } from './cellRange.ts';
import type { Workbook } from './workbook.ts';

/** A named range (`TaxRate`). `scope` null = workbook-global, else the owning sheet's name. */
export interface NamedRange {
  readonly name: string;
  readonly target: CellRange;
  readonly scope?: string | null;
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
