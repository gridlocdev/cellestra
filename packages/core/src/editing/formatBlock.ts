import type { Cell } from '../model/cell.ts';
import { CellRange } from '../model/cellRange.ts';
import { MergedRegion } from '../model/mergedRegion.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import type { CellStyle } from '../styling/cellStyle.ts';
import {
  EditResult,
  requireSheet,
  type ChangedCell,
  type EditCommand,
  type Position,
} from './editCommand.ts';

/** A captured pattern of formats (and merges anchored inside it) for the Format Painter. */
export class FormatBlock {
  readonly rows: number;
  readonly columns: number;
  /** `styles[r][c]`: the effective style of each source cell. */
  readonly styles: readonly (readonly CellStyle[])[];
  /** Merge anchors relative to the block's top-left, paired with `mergeSpans`. */
  readonly mergeAnchors: readonly Position[];
  readonly mergeSpans: readonly (readonly [rows: number, columns: number])[];

  private constructor(
    rows: number,
    columns: number,
    styles: readonly (readonly CellStyle[])[],
    mergeAnchors: readonly Position[],
    mergeSpans: readonly (readonly [number, number])[],
  ) {
    this.rows = rows;
    this.columns = columns;
    this.styles = styles;
    this.mergeAnchors = mergeAnchors;
    this.mergeSpans = mergeSpans;
  }

  static capture(workbook: Workbook, sheet: Worksheet, range: CellRange): FormatBlock {
    const styles: CellStyle[][] = [];
    for (let r = 0; r < range.rowCount; r++) {
      const row: CellStyle[] = [];
      // Effective, not own: the painter picks up how a cell looks, column default included.
      for (let c = 0; c < range.columnCount; c++) {
        row.push(workbook.styles.get(sheet.effectiveStyleId(range.top + r, range.left + c)));
      }
      styles.push(row);
    }
    const anchors: Position[] = [];
    const spans: [number, number][] = [];
    for (const region of sheet.merges) {
      if (range.contains(region.range.top, region.range.left)) {
        anchors.push([region.range.top - range.top, region.range.left - range.left]);
        spans.push([region.range.rowCount, region.range.columnCount]);
      }
    }
    return new FormatBlock(range.rowCount, range.columnCount, styles, anchors, spans);
  }

  /** Paint the pattern over `target`, tiling it, as one "Format Painter" undo step. */
  applyTo(sheet: string, target: CellRange): EditCommand {
    return new PaintCommand(this, sheet, target);
  }

  /** The rectangle a painter drag from the anchor to (row, column) covers. */
  static paintTarget(anchorRow: number, anchorColumn: number, row: number, column: number): CellRange {
    return CellRange.fromBounds(
      Math.min(anchorRow, row),
      Math.min(anchorColumn, column),
      Math.max(anchorRow, row),
      Math.max(anchorColumn, column),
    );
  }
}

class PaintCommand implements EditCommand {
  readonly label = 'Format Painter';
  private readonly block: FormatBlock;
  private readonly sheetName: string;
  private readonly target: CellRange;
  private cellSnapshot: [number, number, Cell | null][] | null = null;
  private mergeSnapshot: MergedRegion[] | null = null;

  constructor(block: FormatBlock, sheetName: string, target: CellRange) {
    this.block = block;
    this.sheetName = sheetName;
    this.target = target;
  }

  apply(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
    const t = this.target;
    this.cellSnapshot = [...t.positions()].map(([r, c]) => [r, c, sheet.find(r, c)?.clone() ?? null]);
    this.mergeSnapshot = [...sheet.merges];
    const changed: ChangedCell[] = [];
    for (let r = 0; r < t.rowCount; r++) {
      for (let c = 0; c < t.columnCount; c++) {
        const row = t.top + r;
        const column = t.left + c;
        const style = this.block.styles[r % this.block.rows]![c % this.block.columns]!;
        const cell = sheet.getOrCreate(row, column);
        cell.styleId = workbook.styles.intern(style);
        sheet.pruneIfBlank(row, column, cell);
        changed.push({ sheet: this.sheetName, row, column });
      }
    }
    this.applyMerges(sheet);
    return EditResult.cells(changed);
  }

  undo(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
    if (!this.cellSnapshot || !this.mergeSnapshot) return EditResult.none;
    for (const [row, column, before] of this.cellSnapshot)
      sheet.putCell(row, column, before?.clone() ?? null);
    sheet.replaceMerges(this.mergeSnapshot);
    return EditResult.cells(
      this.cellSnapshot.map(([row, column]) => ({ sheet: this.sheetName, row, column })),
    );
  }

  private applyMerges(sheet: Worksheet): void {
    const b = this.block;
    const t = this.target;
    if (b.mergeAnchors.length === 0) return;
    // Clear merges overlapping the target so the painted pattern wins (mirrors paste).
    for (const region of [...sheet.merges]) {
      if (region.range.intersects(t)) sheet.removeMerge(region.range.top, region.range.left);
    }
    // Recreate the source merges tiled with the style pattern, clamped to the target.
    const placed: CellRange[] = [];
    for (let tileRow = 0; tileRow < t.rowCount; tileRow += b.rows) {
      for (let tileCol = 0; tileCol < t.columnCount; tileCol += b.columns) {
        b.mergeAnchors.forEach(([ar, ac], i) => {
          const [sr, sc] = b.mergeSpans[i]!;
          const top = t.top + tileRow + ar;
          const left = t.left + tileCol + ac;
          if (top > t.bottom || left > t.right) return;
          const bottom = Math.min(top + sr - 1, t.bottom);
          const right = Math.min(left + sc - 1, t.right);
          if (bottom === top && right === left) return; // clamped to one cell: nothing to merge
          const range = CellRange.fromBounds(top, left, bottom, right);
          if (placed.some((p) => p.intersects(range))) return;
          sheet.addMerge(new MergedRegion(range));
          placed.push(range);
        });
      }
    }
  }
}
