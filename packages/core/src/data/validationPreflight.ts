import type { Worksheet } from '../model/worksheet.ts';
import type { CellValue } from '../values/cellValue.ts';
import { Validator, type RangeResolver } from './dataValidation.ts';

/** A value about to land in a cell (from paste or fill). */
export interface IncomingWrite {
  readonly row: number;
  readonly column: number;
  readonly value: CellValue;
  readonly isFormula: boolean;
}

/** Batch validation pre-flight for paste/fill (PRD §8.13). */
export const ValidationPreflight = {
  /**
   * Positions whose incoming value breaks a *stop* rule. Formulas are skipped (validated on
   * recalc, not entry); warning/information rules never block.
   */
  findStopViolations(
    sheet: Worksheet,
    incoming: Iterable<IncomingWrite>,
    locale: string,
    resolveRange?: RangeResolver | null,
  ): [row: number, column: number][] {
    const out: [number, number][] = [];
    for (const { row, column, value, isFormula } of incoming) {
      if (isFormula) continue;
      const rule = sheet.findValidation(row, column);
      if (!rule || rule.response !== 'stop') continue;
      if (!Validator.isValid(rule, value, locale, resolveRange)) out.push([row, column]);
    }
    return out;
  },
} as const;
