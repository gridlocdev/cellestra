import { CellStyles, NumberFormats, styleKey, type Workbook, type Worksheet } from '@cellestra/core';

/**
 * What saving `workbook` as CSV would throw away. CSV keeps one sheet's displayed values (number
 * formats survive as text); other sheets, formulas, cell formatting and merges do not.
 * Empty when CSV loses nothing worth asking about.
 */
export function csvLosses(workbook: Workbook, sheet: Worksheet): string[] {
  const losses: string[] = [];
  if (workbook.sheetCount > 1)
    losses.push(`${workbook.sheetCount - 1} other sheet(s): only “${sheet.name}” is kept`);
  let formulas = false;
  let styles = false;
  const plain = styleKey(CellStyles.default);
  const styled = new Map<number, boolean>();
  const isStyled = (id: number): boolean => {
    let v = styled.get(id);
    if (v === undefined) {
      v = styleKey({ ...workbook.styles.get(id), numberFormat: NumberFormats.general }) !== plain;
      styled.set(id, v);
    }
    return v;
  };
  const used = sheet.getUsedRange();
  if (used) {
    for (let r = used.top; r <= used.bottom && !(formulas && styles); r++) {
      for (let c = used.left; c <= used.right; c++) {
        const cell = sheet.find(r, c);
        if (!cell) continue;
        if (cell.formula !== null) formulas = true;
        if (cell.styleId !== 0 && isStyled(cell.styleId)) styles = true;
        if (formulas && styles) break;
      }
    }
  }
  if (formulas) losses.push('formulas');
  if (styles) losses.push('formatting');
  if (sheet.merges.length > 0) losses.push('merged cells');
  return losses;
}
