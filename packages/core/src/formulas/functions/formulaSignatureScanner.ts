/** The function call whose parentheses enclose a caret, and which argument the caret is in. */
export interface FormulaCallSite {
  readonly functionName: string;
  readonly argumentIndex: number;
}

const isNameChar = (c: string): boolean => /^[\p{L}\p{Nd}_]$/u.test(c);

/**
 * Finds the function call a caret sits inside while a formula is typed, so the editor can show a
 * positional argument hint. Pure text scanning, so it works on half-written formulas.
 */
export const FormulaSignatureScanner = {
  /**
   * The innermost named call whose parentheses enclose `caret` in `formula` (a leading `=` is fine),
   * with the zero-based argument index; null when the caret is not inside a function call.
   */
  find(formula: string, caret: number): FormulaCallSite | null {
    if (!formula) return null;
    const end = Math.max(0, Math.min(caret, formula.length));
    const stack: { nameStart: number; nameEnd: number; argIndex: number }[] = [];
    let inString = false;
    let braces = 0; // inside an inline array, commas separate its elements, not arguments
    for (let i = 0; i < end; i++) {
      const c = formula[i]!;
      if (inString) {
        if (c === '"') {
          if (formula[i + 1] === '"') i++;
          else inString = false;
        }
        continue;
      }
      if (c === '"') inString = true;
      else if (c === '{') braces++;
      else if (c === '}') braces = Math.max(0, braces - 1);
      else if (c === '(') {
        let j = i - 1;
        while (j >= 0 && isNameChar(formula[j]!)) j--;
        stack.push({ nameStart: j + 1, nameEnd: i, argIndex: 0 });
      } else if (c === ')') stack.pop();
      else if (c === ',' && braces === 0 && stack.length > 0) stack[stack.length - 1]!.argIndex++;
    }
    while (stack.length > 0) {
      const { nameStart, nameEnd, argIndex } = stack.pop()!;
      if (nameEnd > nameStart)
        return { functionName: formula.slice(nameStart, nameEnd), argumentIndex: argIndex };
    }
    return null;
  },
} as const;
