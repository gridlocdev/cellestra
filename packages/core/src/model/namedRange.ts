import type { CellRange } from './cellRange.ts';

/** A named range (`TaxRate`). `scope` null = workbook-global, else the owning sheet's name. */
export interface NamedRange {
  readonly name: string;
  readonly target: CellRange;
  readonly scope?: string | null;
}
