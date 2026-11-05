import type { CellRange } from '../model/cellRange.ts';
import type { Position } from './editCommand.ts';
import { OVERSIZED_CELL_CAP } from './editCommands.ts';

/** Where a block lands: one anchor per tile, whether it tiled, and whether sizes failed to reconcile. */
export interface PastePlan {
  readonly anchors: readonly Position[];
  readonly tiled: boolean;
  readonly shapeMismatch: boolean;
}

/** Excel's paste geometry: expand from one cell, fill from a single cell, tile exact multiples. */
export const PastePlanner = {
  plan(blockRows: number, blockColumns: number, selection: CellRange, isCut: boolean): PastePlan {
    const once: PastePlan = {
      anchors: [[selection.top, selection.left]],
      tiled: false,
      shapeMismatch: false,
    };
    // A move, or a paste onto one active cell: anchor once and expand.
    if (isCut || selection.isSingleCell) return once;
    // A single source cell fills the selection — but never the whole sheet (sparse by design).
    if (blockRows === 1 && blockColumns === 1) {
      if (isOversized(selection)) return once;
      return { anchors: [...selection.positions()], tiled: true, shapeMismatch: false };
    }
    if (
      !isOversized(selection) &&
      selection.rowCount % blockRows === 0 &&
      selection.columnCount % blockColumns === 0
    ) {
      const anchors: Position[] = [];
      for (let r = 0; r < selection.rowCount; r += blockRows) {
        for (let c = 0; c < selection.columnCount; c += blockColumns) {
          anchors.push([selection.top + r, selection.left + c]);
        }
      }
      return { anchors, tiled: anchors.length > 1, shapeMismatch: false };
    }
    return { ...once, shapeMismatch: true };
  },
} as const;

function isOversized(selection: CellRange): boolean {
  return selection.rowCount * selection.columnCount > OVERSIZED_CELL_CAP;
}
