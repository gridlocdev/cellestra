import type { CellRichText, UnmodeledCellContent, UnparsedFormulaSource } from '../model/cell.ts';
import { CellRange } from '../model/cellRange.ts';
import type { Hyperlink } from '../model/hyperlink.ts';
import type { Worksheet } from '../model/worksheet.ts';
import type { CellValue } from '../values/cellValue.ts';
import { clampToUsed, OVERSIZED_CELL_CAP } from './editCommands.ts';

/** One copied cell: value, formula source (no `=`), hyperlink, style id, comment and rich text. */
export interface ClipboardCell {
  readonly value: CellValue;
  readonly formula: string | null;
  /** A loaded formula's file text, kept while the parser cannot read the formula. */
  readonly formulaSource: UnparsedFormulaSource | null;
  readonly styleId: number;
  readonly hyperlink: Hyperlink | null;
  /** What the file gave the cell that the model doesn't model, such as its comment. */
  readonly unmodeled: UnmodeledCellContent | null;
  /** The cell's text as the file marked it up. */
  readonly richText: CellRichText | null;
}

/** A copied or cut rectangle; `cells[r][c]` is null where the source had no cell. */
export class ClipboardBlock {
  readonly sourceSheet: string;
  readonly sourceRange: CellRange;
  readonly isCut: boolean;
  readonly cells: readonly (readonly (ClipboardCell | null)[])[];

  constructor(
    sourceSheet: string,
    sourceRange: CellRange,
    isCut: boolean,
    cells: readonly (readonly (ClipboardCell | null)[])[],
  ) {
    this.sourceSheet = sourceSheet;
    this.sourceRange = sourceRange;
    this.isCut = isCut;
    this.cells = cells;
  }

  get rows(): number {
    return this.cells.length;
  }

  get columns(): number {
    return this.cells[0]?.length ?? 0;
  }

  /** The range a copy actually reads: oversized selections clamp to the used range (else the top-left). */
  static copyRange(sheet: Worksheet, range: CellRange): CellRange {
    return range.rowCount * range.columnCount < OVERSIZED_CELL_CAP
      ? range
      : (clampToUsed(sheet, range) ?? new CellRange(range.topLeft.relative()));
  }

  static capture(sheet: Worksheet, range: CellRange, isCut: boolean): ClipboardBlock {
    const r0 = ClipboardBlock.copyRange(sheet, range);
    const cells: (ClipboardCell | null)[][] = [];
    for (let r = 0; r < r0.rowCount; r++) {
      const row: (ClipboardCell | null)[] = [];
      for (let c = 0; c < r0.columnCount; c++) {
        const cell = sheet.find(r0.top + r, r0.left + c);
        row.push(
          cell
            ? {
                value: cell.value,
                formula: cell.formula,
                formulaSource: cell.formulaSource,
                styleId: cell.styleId,
                hyperlink: cell.hyperlink,
                unmodeled: cell.unmodeled,
                richText: cell.richText,
              }
            : null,
        );
      }
      cells.push(row);
    }
    return new ClipboardBlock(sheet.name, r0, isCut, cells);
  }
}
