import { ReferenceAdjuster } from '../formulas/references/referenceAdjuster.ts';
import { Cell } from '../model/cell.ts';
import type { CellRange } from '../model/cellRange.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { cellKey } from '../model/sheetDimensions.ts';
import { ClipboardBlock, type ClipboardCell } from './clipboardBlock.ts';
import { EditResult, requireSheet, type EditCommand } from './editCommand.ts';
import { rewriteFormula } from './formulaRewrite.ts';
import type { ProjectedWrite } from './pasteBlockCommand.ts';

/**
 * Fill handle: tile the source block over the fill range (the source itself stays as is), shifting
 * relative formula references by each target's offset from its source cell.
 */
export class FillCommand implements EditCommand {
  readonly label = 'Fill';
  private readonly sheetName: string;
  private readonly source: CellRange;
  private readonly fill: CellRange;
  private before: Map<number, [row: number, column: number, before: Cell | null]> | null = null;

  constructor(sheetName: string, source: CellRange, fill: CellRange) {
    this.sheetName = sheetName;
    this.source = source;
    this.fill = fill;
  }

  apply(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
    this.before = new Map();
    for (const t of this.targets(sheet)) {
      this.before.set(cellKey(t.row, t.column), [
        t.row,
        t.column,
        sheet.find(t.row, t.column)?.clone() ?? null,
      ]);
      writeCell(sheet, t.row, t.column, t.cell, t.row - t.sourceRow, t.column - t.sourceColumn);
    }
    return EditResult.cells(this.changedRefs());
  }

  undo(workbook: Workbook): EditResult {
    if (!this.before) return EditResult.none;
    const sheet = workbook.findSheet(this.sheetName);
    for (const [row, column, before] of this.before.values())
      sheet?.putCell(row, column, before?.clone() ?? null);
    return EditResult.cells(this.changedRefs());
  }

  /** The values this fill would write, without applying it. */
  *projectWrites(workbook: Workbook): IterableIterator<ProjectedWrite> {
    const sheet = workbook.findSheet(this.sheetName);
    if (!sheet) return;
    for (const t of this.targets(sheet)) {
      if (t.cell)
        yield { row: t.row, column: t.column, value: t.cell.value, isFormula: t.cell.formula !== null };
    }
  }

  private *targets(sheet: Worksheet) {
    const block = ClipboardBlock.capture(sheet, this.source, false).cells;
    const s = this.source;
    for (const [row, column] of this.fill.positions()) {
      if (s.contains(row, column)) continue;
      const sourceRow = s.top + mod(row - s.top, s.rowCount);
      const sourceColumn = s.left + mod(column - s.left, s.columnCount);
      const cell = block[sourceRow - s.top]?.[sourceColumn - s.left] ?? null;
      yield { row, column, sourceRow, sourceColumn, cell };
    }
  }

  private changedRefs() {
    return [...this.before!.values()].map(([row, column]) => ({ sheet: this.sheetName, row, column }));
  }
}

function writeCell(
  sheet: Worksheet,
  row: number,
  column: number,
  source: ClipboardCell | null,
  rowDelta: number,
  columnDelta: number,
): void {
  if (!source) {
    sheet.putCell(row, column, null);
    return;
  }
  const cell = new Cell();
  cell.value = source.value;
  cell.formula =
    source.formula === null
      ? null
      : rewriteFormula(source.formula, (expr) => ReferenceAdjuster.forCopy(expr, columnDelta, rowDelta));
  cell.formulaSource = source.formulaSource;
  cell.hyperlink = source.hyperlink;
  cell.styleId = source.styleId;
  sheet.putCell(row, column, cell);
}

const mod = (value: number, modulus: number): number => ((value % modulus) + modulus) % modulus;
