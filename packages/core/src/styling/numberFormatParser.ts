export type NumberFormatKind =
  | 'general'
  | 'number'
  | 'percent'
  | 'currency'
  | 'scientific'
  | 'fraction'
  | 'date'
  | 'time'
  | 'dateTime'
  | 'text';

/** `[$` followed by a symbol: `[$€-407]`, `[$USD]`. A bare locale tag `[$-409]` has none. */
const CURRENCY_TOKEN = /\[\$[^\]-]/;

/** Classify a format code by its pattern letters only — literals (`"days"`, `[RED]`, `\d`) don't count. */
export function classifyNumberFormat(code: string | null | undefined): NumberFormatKind {
  if (!code || code.toLowerCase() === 'general') return 'general';
  if (code === '@') return 'text';
  // A currency token (`[$€-407]`, `[$USD]`) is currency by definition, whatever letters its locale
  // tag contains. A locale tag alone (`[$-409]`) only sets the language, so the pattern decides.
  if (CURRENCY_TOKEN.test(code)) return 'currency';
  const pattern = withoutLiterals(code);
  const hasDate = /[yYdD]/.test(pattern);
  const hasMonth = /[mM]/.test(pattern);
  const hasTime = /[hHsS]/.test(pattern);
  if (hasDate && hasTime) return 'dateTime';
  if (hasDate || (hasMonth && !hasTime)) return 'date';
  if (hasTime) return 'time';
  if (pattern.includes('%')) return 'percent';
  if (/[$€£¥¤]/.test(pattern)) return 'currency';
  if (/e/i.test(pattern) && /[0#]/.test(pattern)) return 'scientific';
  // `# ?/?`, `# ??/??`, `?/4`: a placeholder, a slash, then a placeholder or a fixed denominator.
  if (/[0#?]\s*\/\s*[0-9#?]/.test(pattern)) return 'fraction';
  if (/[0#]/.test(pattern)) return 'number';
  return 'general';
}

export function isDateLikeKind(kind: NumberFormatKind): boolean {
  return kind === 'date' || kind === 'time' || kind === 'dateTime';
}

/** Strip quoted literals, bracketed tokens (except elapsed-time `[h]`/`[mm]`/`[ss]`), and escapes. */
export function withoutLiterals(code: string): string {
  if (!/["[\\]/.test(code)) return code;
  let kept = '';
  for (let i = 0; i < code.length; i++) {
    const c = code[i]!;
    if (c === '"') {
      i = skip(code, i + 1, '"');
    } else if (c === '[') {
      const end = skip(code, i + 1, ']');
      const inside = code.slice(i + 1, Math.max(i + 1, end));
      if (inside.length > 0 && /^[hHmMsS]+$/.test(inside)) kept += inside;
      i = end;
    } else if (c === '\\') {
      i++;
    } else kept += c;
  }
  return kept;
}

function skip(code: string, from: number, close: string): number {
  const end = code.indexOf(close, from);
  return end < 0 ? code.length - 1 : end;
}
