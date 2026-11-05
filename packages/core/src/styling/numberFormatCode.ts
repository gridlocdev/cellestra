import { classifyNumberFormat } from './numberFormatParser.ts';

/** Increase/Decrease Decimals: rewrite the first numeric run's fraction, keep everything else. */
export const NumberFormatCode = {
  increaseDecimals: (code: string | null | undefined): string => adjustDecimals(code, +1),
  decreaseDecimals: (code: string | null | undefined): string => adjustDecimals(code, -1),
  adjustDecimals,
} as const;

function adjustDecimals(code: string | null | undefined, delta: number): string {
  // Seed General from a grouped number so stepping up gives "#,##0.0" and down stays "#,##0".
  const src = !code || code.toLowerCase() === 'general' ? '#,##0' : code;
  const kind = classifyNumberFormat(src);
  if (kind === 'date' || kind === 'time' || kind === 'dateTime' || kind === 'text') return src;
  let start = -1;
  let end = -1;
  for (let i = 0; i < src.length; i++) {
    const isMantissa = '#0,.'.includes(src[i]!);
    if (isMantissa && start < 0) start = i;
    if (isMantissa) end = i;
    else if (start >= 0) break; // ignore later numeric chars (the "00" of E+00)
  }
  if (start < 0) return src;
  const prefix = src.slice(0, start);
  const mantissa = src.slice(start, end + 1);
  const suffix = src.slice(end + 1);
  const dot = mantissa.indexOf('.');
  const integerPart = dot < 0 ? mantissa : mantissa.slice(0, dot);
  const fractionDigits = dot < 0 ? 0 : mantissa.length - dot - 1;
  const target = Math.max(0, fractionDigits + delta);
  return prefix + integerPart + (target > 0 ? '.' + '0'.repeat(target) : '') + suffix;
}
