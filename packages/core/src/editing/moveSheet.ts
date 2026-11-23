import type { Expr } from '../formulas/ast/expr.ts';
import { FormulaWriter } from '../formulas/formulaWriter.ts';
import { Parser } from '../formulas/parsing/parser.ts';
import { ReferenceAdjuster } from '../formulas/references/referenceAdjuster.ts';
import type { NamedRange } from '../model/namedRange.ts';
import type { Workbook } from '../model/workbook.ts';

/**
 * Move the sheet at tab `from` to tab `to`. A 3D range (`Sheet1.A1:Sheet3.B2`) keeps its two end
 * sheets and spans whatever lies between them now; when the move puts its last sheet before its
 * first, its text swaps them so it reads in tab order, as LibreOffice does. Cell formulas and names
 * that hold an expression are rewritten.
 */
export function moveSheet(workbook: Workbook, from: number, to: number): void {
  workbook.moveSheet(from, to);
  const order = new Map(workbook.worksheets.map((s, i) => [s.name.toLowerCase(), i]));
  const position = (sheet: string): number => order.get(sheet.toLowerCase()) ?? -1;
  const transform = (expr: Expr): Expr => ReferenceAdjuster.forSheetOrder(expr, position);
  // Null when nothing changes, so formulas without a reversed 3D range keep their text exactly.
  const reordered = (formula: string): string | null => {
    if (!formula.includes(':')) return null;
    const parsed = Parser.tryParse(formula);
    if (!parsed.ok) return null;
    const after = FormulaWriter.write(transform(parsed.expr));
    return after === FormulaWriter.write(parsed.expr) ? null : after;
  };
  for (const sheet of workbook.worksheets) {
    for (const [, , cell] of sheet.enumerateCells()) {
      const formula = cell.formula === null ? null : reordered(cell.formula);
      if (formula !== null) cell.formula = formula;
    }
  }
  workbook.replaceNamedRanges(
    workbook.namedRanges.map((n): NamedRange => {
      if (n.expression === undefined) return n;
      const expression = reordered(n.expression);
      return expression === null ? n : { ...n, expression };
    }),
  );
}
