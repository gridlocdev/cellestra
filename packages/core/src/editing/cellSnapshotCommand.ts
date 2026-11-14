import type { Cell } from '../model/cell.ts';
import type { CellRange } from '../model/cellRange.ts';
import { cellKey } from '../model/sheetDimensions.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import {
  EditResult,
  requireSheet,
  type ChangedCell,
  type EditCommand,
  type Position,
} from './editCommand.ts';

type Snapshot = (readonly [row: number, column: number, before: Cell | null])[];

/**
 * The general-purpose cell edit: snapshot the cells at some positions, run a mutation, and undo by
 * putting the snapshots back. Positions may be computed per apply (from the sheet as it is then).
 *
 * `overRange` snapshots a whole range instead: only the cells stored in it, so a mostly empty range
 * costs what it holds, not its area. Undo clears the range and puts those cells back.
 */
export class CellSnapshotCommand implements EditCommand {
  readonly label: string;
  private readonly sheetName: string;
  private positions: readonly Position[];
  private readonly positionFactory: ((sheet: Worksheet) => readonly Position[]) | null;
  private readonly rangeFactory: ((sheet: Worksheet) => CellRange | null) | null = null;
  private range: CellRange | null = null;
  private readonly mutate: (sheet: Worksheet) => void;
  private snapshot: Snapshot | null = null;

  constructor(
    label: string,
    sheetName: string,
    positions:
      | readonly Position[]
      | ((sheet: Worksheet) => readonly Position[])
      | { readonly range: (sheet: Worksheet) => CellRange | null },
    mutate: (sheet: Worksheet) => void,
  ) {
    this.label = label;
    this.sheetName = sheetName;
    this.mutate = mutate;
    this.positions = [];
    this.positionFactory = null;
    if (typeof positions === 'function') this.positionFactory = positions;
    else if ('range' in positions) this.rangeFactory = positions.range;
    else this.positions = positions;
  }

  /** Snapshot every cell stored in `range` (computed per apply; null for nothing), then `mutate`. */
  static overRange(
    label: string,
    sheetName: string,
    range: (sheet: Worksheet) => CellRange | null,
    mutate: (sheet: Worksheet) => void,
  ): CellSnapshotCommand {
    return new CellSnapshotCommand(label, sheetName, { range }, mutate);
  }

  apply(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
    if (this.rangeFactory) return this.applyOverRange(sheet);
    if (this.positionFactory) this.positions = this.positionFactory(sheet);
    this.snapshot = this.positions.map(([row, column]) => [
      row,
      column,
      sheet.find(row, column)?.clone() ?? null,
    ]);
    this.mutate(sheet);
    return EditResult.cells(this.refs());
  }

  undo(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
    if (!this.snapshot) return EditResult.none;
    if (this.range) {
      // The cells the edit left in the range go, then the ones it found there come back.
      const after = sheet.cellsIn(this.range);
      for (const [row, column] of after) sheet.putCell(row, column, null);
      for (const [row, column, before] of this.snapshot) sheet.putCell(row, column, before?.clone() ?? null);
      this.positions = changedPositions(this.snapshot, after);
      return EditResult.cells(this.refs());
    }
    for (const [row, column, before] of this.snapshot) sheet.putCell(row, column, before?.clone() ?? null);
    return EditResult.cells(this.refs());
  }

  private applyOverRange(sheet: Worksheet): EditResult {
    this.range = this.rangeFactory!(sheet);
    if (!this.range) {
      this.snapshot = [];
      this.positions = [];
      this.mutate(sheet);
      return EditResult.none;
    }
    this.snapshot = sheet.cellsIn(this.range).map(([row, column, cell]) => [row, column, cell.clone()]);
    this.mutate(sheet);
    // What changed: every position that held a cell before or holds one now.
    this.positions = changedPositions(this.snapshot, sheet.cellsIn(this.range));
    return EditResult.cells(this.refs());
  }

  private refs(): ChangedCell[] {
    return this.positions.map(([row, column]) => ({ sheet: this.sheetName, row, column }));
  }
}

/** The positions in `before` and `after`, each once. */
function changedPositions(
  before: Snapshot,
  after: readonly (readonly [row: number, column: number, ...rest: unknown[]])[],
): Position[] {
  const seen = new Set<number>();
  const out: Position[] = [];
  for (const [row, column] of [...before, ...after]) {
    const key = cellKey(row, column);
    if (seen.has(key)) continue;
    seen.add(key);
    out.push([row, column]);
  }
  return out;
}
