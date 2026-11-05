import type { Cell } from '../model/cell.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import type { ChangedCell } from './editCommand.ts';

/** Before-images of cells across sheets, first capture wins — the undo half of multi-sheet edits. */
export class CellSnapshots {
  private readonly map = new Map<string, { ref: ChangedCell; before: Cell | null }>();

  capture(sheet: Worksheet, row: number, column: number): void {
    const key = `${sheet.name}\u0000${row}\u0000${column}`;
    if (!this.map.has(key)) {
      this.map.set(key, {
        ref: { sheet: sheet.name, row, column },
        before: sheet.find(row, column)?.clone() ?? null,
      });
    }
  }

  restore(workbook: Workbook): void {
    for (const { ref, before } of this.map.values()) {
      workbook.findSheet(ref.sheet)?.putCell(ref.row, ref.column, before?.clone() ?? null);
    }
  }

  refs(): ChangedCell[] {
    return [...this.map.values()].map((s) => s.ref);
  }
}
