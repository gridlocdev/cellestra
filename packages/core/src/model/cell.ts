import { CellValue } from '../values/cellValue.ts';
import type { Hyperlink } from './hyperlink.ts';

/**
 * One cell's content. A literal cell has `formula === null` and `value` is the content; a formula
 * cell holds the source without its leading `=` and `value` is the cached/computed result — shown
 * immediately on open, then refreshed by recalculation (PRD §8.6).
 */
export class Cell {
  formula: string | null = null;
  value: CellValue = CellValue.empty;
  /** Index into the workbook's style pool; 0 is the default style. */
  styleId = 0;
  hyperlink: Hyperlink | null = null;

  get isFormula(): boolean {
    return this.formula !== null;
  }

  /** Carries nothing worth storing — used to prune sparse storage. */
  get isBlank(): boolean {
    return this.formula === null && this.value.kind === 'empty' && this.styleId === 0 && this.hyperlink === null;
  }

  clone(): Cell {
    const c = new Cell();
    c.formula = this.formula;
    c.value = this.value;
    c.styleId = this.styleId;
    c.hyperlink = this.hyperlink;
    return c;
  }
}
