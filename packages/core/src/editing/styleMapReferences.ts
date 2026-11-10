import type { Expr } from '../formulas/ast/expr.ts';
import { FormulaWriter } from '../formulas/formulaWriter.ts';
import { OpenFormulaTranslator } from '../formulas/openFormulaTranslator.ts';
import { Parser } from '../formulas/parsing/parser.ts';
import type { Workbook } from '../model/workbook.ts';
import type { CellStyle, CellStyleMap } from '../styling/cellStyle.ts';
import type { ReferenceTransform } from './workbookReferences.ts';

/** Each pooled style that carries conditional maps, by id, as it was. */
export type StyleMapSnapshot = readonly (readonly [number, CellStyle])[];

/** A bracketed OpenFormula reference inside a `style:condition` (`[.A1]`, `[Sheet2.A1:.B3]`). */
const BRACKETED = /\[[^\]]*\]/g;
/** A home no sheet has: with no base cell, only sheet-qualified references can be adjusted. */
const NO_HOME = '\u0000';

export function captureStyleMaps(workbook: Workbook): StyleMapSnapshot {
  const out: [number, CellStyle][] = [];
  workbook.styles.styles.forEach((style, id) => {
    if (style.conditionalMaps?.length) out.push([id, style]);
  });
  return out;
}

export function restoreStyleMaps(workbook: Workbook, snapshot: StyleMapSnapshot): void {
  for (const [id, style] of snapshot) workbook.styles.replace(id, style);
}

/**
 * Move conditional maps' base cells and the references in their conditions the way formulas move.
 * The style keeps its id, so every cell using it follows. A base cell that was deleted is left as
 * it was (the condition's relative references still need some anchor); a deleted reference in the
 * condition becomes `#REF!`, as in a formula.
 */
export function transformStyleMaps(workbook: Workbook, transform: ReferenceTransform): void {
  const pool = workbook.styles;
  pool.styles.forEach((style, id) => {
    const maps = style.conditionalMaps;
    if (!maps?.length) return;
    let changed = false;
    const next = maps.map((map) => {
      const moved = transformMap(map, transform);
      if (moved !== map) changed = true;
      return moved;
    });
    if (changed) pool.replace(id, { ...style, conditionalMaps: next });
  });
}

function transformMap(map: CellStyleMap, transform: ReferenceTransform): CellStyleMap {
  const base = map.baseCellAddress === null ? null : parseReference(`[${map.baseCellAddress}]`);
  const home = (base?.kind === 'reference' ? base.address.sheet : null) ?? NO_HOME;
  let baseCellAddress = map.baseCellAddress;
  if (base) {
    const moved = transform(base, home);
    if (moved.kind === 'reference') {
      const text = writeReference(moved);
      if (text !== writeReference(base)) baseCellAddress = text.slice(1, -1);
    }
  }
  const condition = map.condition.replace(BRACKETED, (token) => {
    const expr = parseReference(token);
    if (!expr) return token;
    const text = writeReference(transform(expr, home));
    // Unchanged references keep their exact spelling ($ flags, sheet quoting).
    return text === writeReference(expr) ? token : text;
  });
  if (condition === map.condition && baseCellAddress === map.baseCellAddress) return map;
  return { ...map, condition, baseCellAddress };
}

/** An OpenFormula reference or range, or null when `odf` is something else. */
function parseReference(odf: string): Expr | null {
  const parsed = Parser.tryParse(OpenFormulaTranslator.fromOpenFormula(`of:=${odf}`));
  return parsed.ok && (parsed.expr.kind === 'reference' || parsed.expr.kind === 'range') ? parsed.expr : null;
}

function writeReference(expr: Expr): string {
  return OpenFormulaTranslator.toOpenFormula(FormulaWriter.write(expr)).slice('of:='.length);
}
