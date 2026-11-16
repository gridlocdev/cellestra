import type { Expr } from '../ast/expr.ts';

/** Whether `expr` can stand for cells: a reference, range, name, broken reference, reference operator or call. */
export function isReferenceExpr(expr: Expr): boolean {
  switch (expr.kind) {
    case 'reference':
    case 'range':
    case 'sheetSpan':
    case 'name':
    case 'refError':
      return true;
    case 'binary':
      return expr.operator === 'union' || expr.operator === 'intersect';
    case 'call':
      return true; // a reference-returning function (INDEX); any other is #VALUE! when evaluated
    case 'array':
    case 'literal':
    case 'unary':
    default:
      return false;
  }
}

/** How deep `expr` nests, counted without recursion (long operator chains build deep trees). */
export function treeDepth(expr: Expr): number {
  let deepest = 0;
  const stack: [Expr, number][] = [[expr, 1]];
  for (let item = stack.pop(); item; item = stack.pop()) {
    const [node, depth] = item;
    if (depth > deepest) deepest = depth;
    if (node.kind === 'unary') stack.push([node.operand, depth + 1]);
    else if (node.kind === 'binary') stack.push([node.left, depth + 1], [node.right, depth + 1]);
    else if (node.kind === 'call') for (const a of node.args) stack.push([a, depth + 1]);
  }
  return deepest;
}
