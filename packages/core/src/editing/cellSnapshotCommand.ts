import type { Cell } from '../model/cell.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import {
  EditResult,
  requireSheet,
  type ChangedCell,
  type EditCommand,
  type Position,
} from './editCommand.ts';

/**
 * The general-purpose cell edit: snapshot the cells at some positions, run a mutation, and undo by
 * putting the snapshots back. Positions may be computed per apply (from the sheet as it is then).
 */
export class CellSnapshotCommand implements EditCommand {
  readonly label: string;
  private readonly sheetName: string;
  private positions: readonly Position[];
  private readonly positionFactory: ((sheet: Worksheet) => readonly Position[]) | null;
  private readonly mutate: (sheet: Worksheet) => void;
  private snapshot: (readonly [row: number, column: number, before: Cell | null])[] | null = null;

  constructor(
    label: string,
    sheetName: string,
    positions: readonly Position[] | ((sheet: Worksheet) => readonly Position[]),
    mutate: (sheet: Worksheet) => void,
  ) {
    this.label = label;
    this.sheetName = sheetName;
    this.mutate = mutate;
    if (typeof positions === 'function') {
      this.positions = [];
      this.positionFactory = positions;
    } else {
      this.positions = positions;
      this.positionFactory = null;
    }
  }

  apply(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
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
    for (const [row, column, before] of this.snapshot) sheet.putCell(row, column, before?.clone() ?? null);
    return EditResult.cells(this.refs());
  }

  private refs(): ChangedCell[] {
    return this.positions.map(([row, column]) => ({ sheet: this.sheetName, row, column }));
  }
}
