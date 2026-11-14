import { CellValue } from '../values/cellValue.ts';
import type { XmlElement } from '../xml/index.ts';
import type { Hyperlink } from './hyperlink.ts';

/**
 * Children a file gave a cell that the model has no field for (a comment, a shape anchored to the
 * cell), split around its text and kept so a save writes them back. The elements are detached and
 * never changed, so clones of a cell share them.
 */
export interface UnmodeledCellContent {
  readonly before: readonly XmlElement[];
  readonly after: readonly XmlElement[];
}

/**
 * A cell's text as a file marked it up (spans with their text styles, links, line breaks), kept so
 * a save writes it back. It stands for the cell only while the cell's text and hyperlink are still
 * the `text` and `hyperlink` it produced; once either is edited the text is written plain.
 */
export interface CellRichText {
  readonly text: string;
  readonly hyperlink: string | null;
  /** The cell's `text:p` elements, detached and never changed, so clones share them. */
  readonly paragraphs: readonly XmlElement[];
}

/**
 * The OpenFormula text a file gave a formula the parser could not read, kept so a save writes it
 * back exactly. It stands for the cell only while the cell's formula is still `formula`.
 */
export interface UnparsedFormulaSource {
  readonly formula: string;
  readonly openFormula: string;
}

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
  /** Content read from a file that the model does not model; travels with the cell. */
  unmodeled: UnmodeledCellContent | null = null;
  /** The marked-up form of the cell's text, read from a file. */
  richText: CellRichText | null = null;
  /** The file's text of an unparsed formula, read from a file. */
  formulaSource: UnparsedFormulaSource | null = null;

  get isFormula(): boolean {
    return this.formula !== null;
  }

  /** Carries nothing worth storing — used to prune sparse storage. */
  get isBlank(): boolean {
    return (
      this.formula === null &&
      this.value.kind === 'empty' &&
      this.styleId === 0 &&
      this.hyperlink === null &&
      this.unmodeled === null
    );
  }

  clone(): Cell {
    const c = new Cell();
    c.formula = this.formula;
    c.value = this.value;
    c.styleId = this.styleId;
    c.hyperlink = this.hyperlink;
    c.unmodeled = this.unmodeled;
    c.richText = this.richText;
    c.formulaSource = this.formulaSource;
    return c;
  }
}
