import { DuplicateRemover, TextToColumns } from '../data/dataTools.ts';
import { CellRange } from '../model/cellRange.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
import type { Worksheet } from '../model/worksheet.ts';
import type { StylePool } from '../styling/stylePool.ts';
import { CellSnapshotCommand } from './cellSnapshotCommand.ts';
import { ActionCommand, type EditCommand } from './editCommand.ts';

/** Undoable data tools that rewrite a block of cells: remove duplicates, text to columns. */
export const DataToolCommands = {
  /** The command, plus a getter for how many rows its last apply removed. */
  removeDuplicates(
    sheet: string,
    range: CellRange,
    keyColumns: readonly number[],
    hasHeader: boolean,
  ): { readonly command: EditCommand; readonly removedCount: () => number } {
    let removed = 0;
    // Cut off below and right of the used range, so a whole-column selection doesn't snapshot a
    // million rows. The top row stays: it may be the header.
    const trimmed = (ws: Worksheet): CellRange | null => {
      const used = ws.getUsedRange();
      if (!used) return null;
      const bottom = Math.min(range.bottom, used.bottom);
      const right = Math.min(range.right, used.right);
      return bottom < range.top || right < range.left
        ? null
        : CellRange.fromBounds(range.top, range.left, bottom, right);
    };
    const command = new CellSnapshotCommand(
      'Remove Duplicates',
      sheet,
      (ws) => [...(trimmed(ws)?.positions() ?? [])],
      (ws) => {
        const r = trimmed(ws);
        removed = r ? DuplicateRemover.remove(ws, r, keyColumns, hasHeader) : 0;
      },
    );
    return { command, removedCount: () => removed };
  },

  /**
   * Split a column's text; snapshots a block as wide as the widest split, so it fully undoes. Only
   * the used rows of a whole-column selection are read. Fields are read as typed input in `locale`
   * (numbers, dates, TRUE/FALSE; by default the document locale), and a date gets its date format.
   */
  textToColumns(
    worksheet: Worksheet,
    sourceColumn: CellRange,
    delimiter: string,
    locale?: string,
  ): EditCommand {
    const source = (ws: Worksheet): CellRange | null => {
      const bottom = Math.min(sourceColumn.bottom, ws.cellExtent()?.lastRow ?? -1);
      return bottom < sourceColumn.top
        ? null
        : CellRange.fromBounds(sourceColumn.top, sourceColumn.left, bottom, sourceColumn.left);
    };
    const block = (ws: Worksheet): CellRange | null => {
      const src = source(ws);
      if (!src) return null;
      let maxFields = 1;
      for (let row = src.top; row <= src.bottom; row++) {
        const v = ws.getValue(row, src.left);
        if (v.kind === 'text')
          maxFields = Math.max(maxFields, TextToColumns.fields(v.value, delimiter).length);
      }
      const right = Math.min(src.left + maxFields - 1, SheetDimensions.maxColumns - 1);
      return CellRange.fromBounds(src.top, src.left, src.bottom, right);
    };
    let styles: StylePool | null = null;
    const split = new CellSnapshotCommand(
      'Text to Columns',
      worksheet.name,
      (ws) => [...(block(ws)?.positions() ?? [])],
      (ws) => {
        const src = source(ws);
        if (src) TextToColumns.split(ws, src, delimiter, { locale: documentLocale, styles });
      },
    );
    let documentLocale = locale;
    return new ActionCommand(
      'Text to Columns',
      (wb) => {
        styles = wb.styles;
        documentLocale = locale ?? wb.settings.effectiveLocale;
        return split.apply(wb);
      },
      (wb) => split.undo(wb),
    );
  },
} as const;
