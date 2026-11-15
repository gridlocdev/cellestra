import type { DataValidation } from '../data/types.ts';
import type { CellRange } from '../model/cellRange.ts';
import { ActionCommand, EditResult, requireSheet, type EditCommand } from './editCommand.ts';

/** Undoable data validation edits. */
export const ValidationCommands = {
  setValidation: (sheet: string, validation: DataValidation): EditCommand =>
    new ActionCommand(
      'Data Validation',
      (wb) => {
        requireSheet(wb, sheet).addValidation(validation);
        return EditResult.none;
      },
      (wb) => {
        requireSheet(wb, sheet).removeValidation(validation);
        return EditResult.none;
      },
    ),

  /**
   * Make `rule` the only rule over `range` (null leaves none there). Rules that reach outside the
   * range keep those parts; the cells inside it lose them. One undo step.
   */
  replaceValidation(sheet: string, range: CellRange, rule: DataValidation | null): EditCommand {
    let removed: DataValidation[] = [];
    let kept: DataValidation[] = [];
    return new ActionCommand(
      'Data Validation',
      (wb) => {
        const ws = requireSheet(wb, sheet);
        removed = ws.validations.filter((v) => v.target.intersects(range));
        kept = removed.flatMap((v) => v.target.subtract(range).map((target) => ({ ...v, target })));
        for (const v of removed) ws.removeValidation(v);
        for (const v of kept) ws.addValidation(v);
        if (rule) ws.addValidation(rule);
        return EditResult.none;
      },
      (wb) => {
        const ws = requireSheet(wb, sheet);
        if (rule) ws.removeValidation(rule);
        for (const v of kept) ws.removeValidation(v);
        for (const v of removed) ws.addValidation(v);
        return EditResult.none;
      },
    );
  },
} as const;
