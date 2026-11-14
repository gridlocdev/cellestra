import type { Cell } from '../model/cell.ts';
import { CellRange } from '../model/cellRange.ts';
import { MergedRegion } from '../model/mergedRegion.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import type { CellStyle } from '../styling/cellStyle.ts';
import {
  combineResults,
  CompositeEditCommand,
  EditResult,
  requireSheet,
  type ChangedCell,
  type EditCommand,
  type Position,
} from './editCommand.ts';
import {
  clampToUsed,
  coversWholeColumns,
  coversWholeRows,
  lineStyleCommand,
  OVERSIZED_CELL_CAP,
} from './editCommands.ts';

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

  /**
   * Paint the pattern over `target`, tiling it, as one "Format Painter" undo step. Whole columns or
   * rows painted with a one-line pattern take it as line styles; any other oversized target is
   * painted only where the sheet is used.
   */
  applyTo(sheet: string, target: CellRange): EditCommand {
    return this.lineCommand(sheet, target) ?? new PaintCommand(this, sheet, target);
  }

  private lineCommand(sheet: string, target: CellRange): EditCommand | null {
    const columns = coversWholeColumns(target);
    const rows = coversWholeRows(target);
    if ((!columns && !rows) || this.mergeAnchors.length > 0) return null;
    const paint = (range: CellRange, style: CellStyle) =>
      lineStyleCommand(sheet, range, () => style, 'Format Painter');
    if (this.rows === 1 && this.columns === 1) return paint(target, this.styles[0]![0]!);
    const steps: EditCommand[] = [];
    if (columns && !rows && this.rows === 1 && target.columnCount <= MAX_LINE_STEPS) {
      for (let c = target.left; c <= target.right; c++)
        steps.push(
          paint(
            CellRange.fromBounds(target.top, c, target.bottom, c),
            this.styles[0]![(c - target.left) % this.columns]!,
          ),
        );
    } else if (rows && !columns && this.columns === 1 && target.rowCount <= MAX_LINE_STEPS) {
      for (let r = target.top; r <= target.bottom; r++)
        steps.push(
          paint(
            CellRange.fromBounds(r, target.left, r, target.right),
            this.styles[(r - target.top) % this.rows]![0]!,
          ),
        );
    } else return null;
    return new CompositeEditCommand('Format Painter', steps);
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

/** The most lines a painter pattern is laid onto one by one, before it paints the used area only. */
const MAX_LINE_STEPS = 1024;

class PaintCommand implements EditCommand {
  readonly label = 'Format Painter';
  private readonly block: FormatBlock;
  private readonly sheetName: string;
  private readonly target: CellRange;
  private cellSnapshot: [number, number, Cell | null][] | null = null;
  private mergeSnapshot: MergedRegion[] | null = null;
  /** The cells the last apply painted. */
  private area: CellRange | null = null;

  constructor(block: FormatBlock, sheetName: string, target: CellRange) {
    this.block = block;
    this.sheetName = sheetName;
    this.target = target;
  }

  apply(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
    const t = this.target;
    // The tiling stays anchored on the target; only the painted area shrinks to the used range.
    const area =
      t.cellCount < OVERSIZED_CELL_CAP ? t : (clampToUsed(sheet, t) ?? new CellRange(t.topLeft.relative()));
    this.cellSnapshot = [...area.positions()].map(([r, c]) => [r, c, sheet.find(r, c)?.clone() ?? null]);
    this.mergeSnapshot = [...sheet.merges];
    const changed: ChangedCell[] = [];
    for (const [row, column] of area.positions()) {
      const style =
        this.block.styles[(row - t.top) % this.block.rows]![(column - t.left) % this.block.columns]!;
      const cell = sheet.getOrCreate(row, column);
      cell.styleId = workbook.styles.intern(style);
      sheet.pruneIfBlank(row, column, cell);
      changed.push({ sheet: this.sheetName, row, column });
    }
    this.applyMerges(sheet, area);
    this.area = area;
    return combineResults([EditResult.cells(changed), this.mergedRows(sheet.merges, area)]);
  }

  /**
   * The rows of the merges over `area`, before the paint and `others` (after it, or undone): a merge
   * reaching past the painted cells changed how those cells look too.
   */
  private mergedRows(others: readonly MergedRegion[], area: CellRange): EditResult {
    const merges = [...(this.mergeSnapshot ?? []), ...others].filter((m) => m.range.intersects(area));
    return combineResults(merges.map((m) => EditResult.rows(this.sheetName, m.range.top, m.range.bottom)));
  }

  undo(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
    if (!this.cellSnapshot || !this.mergeSnapshot) return EditResult.none;
    for (const [row, column, before] of this.cellSnapshot)
      sheet.putCell(row, column, before?.clone() ?? null);
    const painted = [...sheet.merges];
    sheet.replaceMerges(this.mergeSnapshot);
    const cells = EditResult.cells(
      this.cellSnapshot.map(([row, column]) => ({ sheet: this.sheetName, row, column })),
    );
    return this.area ? combineResults([cells, this.mergedRows(painted, this.area)]) : cells;
  }

  private applyMerges(sheet: Worksheet, area: CellRange): void {
    const b = this.block;
    const t = area;
    const origin = this.target;
    if (b.mergeAnchors.length === 0) return;
    // Clear merges overlapping the target so the painted pattern wins (mirrors paste).
    for (const region of [...sheet.merges]) {
      if (region.range.intersects(t)) sheet.removeMerge(region.range.top, region.range.left);
    }
    // Recreate the source merges tiled with the style pattern (anchored on the target), clamped to
    // the painted area.
    const placed: CellRange[] = [];
    const firstRow = t.top - ((t.top - origin.top) % b.rows);
    const firstColumn = t.left - ((t.left - origin.left) % b.columns);
    for (let tileRow = firstRow - t.top; tileRow < t.rowCount; tileRow += b.rows) {
      for (let tileCol = firstColumn - t.left; tileCol < t.columnCount; tileCol += b.columns) {
        b.mergeAnchors.forEach(([ar, ac], i) => {
          const [sr, sc] = b.mergeSpans[i]!;
          const top = t.top + tileRow + ar;
          const left = t.left + tileCol + ac;
          if (top < t.top || left < t.left || top > t.bottom || left > t.right) return;
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
