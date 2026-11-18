import { dataValidation, type DataValidation, type ValidationResponse } from '../../data/types.ts';
import { CellRange } from '../../model/cellRange.ts';
import { excelToInternal } from './xlsxFormulas.ts';

/** One `<dataValidation>` as the file has it, from the worksheet or its `x14` extension. */
export interface XlsxValidation {
  readonly type: string | null;
  readonly sqref: string;
  readonly formula1: string;
  readonly errorStyle: string | null;
  readonly showErrorMessage: boolean;
  readonly showInputMessage: boolean;
  readonly error: string | null;
  readonly prompt: string | null;
}

const RESPONSES: Readonly<Record<string, ValidationResponse>> = {
  stop: 'stop',
  warning: 'warning',
  information: 'information',
};

const NAME = /^[\p{L}_\\][\p{L}\p{Nd}_.]*$/u;

/** `"Yes,No"` (Excel's typed list, quotes doubled inside) → its items, or null when it isn't one. */
function typedItems(formula: string): string[] | null {
  if (formula.length < 2 || !formula.startsWith('"') || !formula.endsWith('"')) return null;
  const items = formula
    .slice(1, -1)
    .replaceAll('""', '"')
    .split(',')
    .map((item) => item.trim())
    .filter((item) => item !== '');
  return items.length > 0 ? items : null;
}

/**
 * A dropdown list rule as Cellestra's rules, one per range it covers (PRD §8.19), or null when
 * Cellestra can't hold it: another kind of rule, or a list from a formula. The list is typed items
 * or a range, given directly or by a named range that `namedRange` finds.
 */
export function importListValidation(
  rule: XlsxValidation,
  namedRange: (name: string) => CellRange | null,
): DataValidation[] | null {
  if (rule.type !== 'list') return null;
  const formula = rule.formula1.trim().replace(/^=/, '');
  const listValues = typedItems(formula);
  let listSource: CellRange | null = null;
  if (listValues === null) {
    listSource = CellRange.tryParse(excelToInternal(formula));
    if (listSource === null && NAME.test(formula)) listSource = namedRange(formula);
    if (listSource === null) return null;
  }
  const targets = rule.sqref
    .split(/\s+/)
    .filter((ref) => ref !== '')
    .map((ref) => CellRange.tryParse(ref));
  if (targets.length === 0 || targets.some((t) => t === null)) return null;
  // Excel lets any entry through when the rule shows no error alert.
  const response = rule.showErrorMessage ? (RESPONSES[rule.errorStyle ?? 'stop'] ?? 'stop') : 'none';
  return targets.map((target) =>
    dataValidation('list', target!, {
      ...(listValues ? { listValues } : { listSource }),
      inputMessage: rule.showInputMessage && rule.prompt ? rule.prompt : null,
      errorMessage: rule.error || null,
      response,
    }),
  );
}
