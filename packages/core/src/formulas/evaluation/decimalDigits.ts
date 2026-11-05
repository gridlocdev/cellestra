/**
 * Decimal digit-string helpers mirroring .NET custom numeric formatting of doubles: the value is
 * first rounded to 15 significant digits, then rounded half-away-from-zero to the requested places.
 * Shared by `Coercion.numberToText` and the `NumberFormatter`.
 */

/** 15 significant digits and the decimal exponent of the first one (`abs` > 0, finite). */
export function significantDigits(abs: number): { digits: string; exponent: number } {
  const s = abs.toExponential(14);
  const e = s.indexOf('e');
  return { digits: s.slice(0, e).replace('.', ''), exponent: Number(s.slice(e + 1)) };
}

/**
 * `abs` (≥ 0) rounded to `decimals` places: integer digits (no leading zeros; "0" for none) and
 * exactly `decimals` fraction digits.
 */
export function fixedDigits(abs: number, decimals: number): { integer: string; fraction: string } {
  if (abs === 0 || !Number.isFinite(abs)) return { integer: '0', fraction: '0'.repeat(decimals) };
  const { digits, exponent } = significantDigits(abs);
  const kept = exponent + 1 + decimals; // digits from the leading one down to the 10^-decimals place
  let n: bigint;
  if (kept < 0) n = 0n;
  else if (kept >= digits.length) n = BigInt(digits) * 10n ** BigInt(kept - digits.length);
  else {
    n = kept === 0 ? 0n : BigInt(digits.slice(0, kept));
    if (digits[kept]! >= '5') n += 1n;
  }
  const text = n.toString().padStart(decimals + 1, '0');
  const integer = text.slice(0, text.length - decimals).replace(/^0+(?=\d)/, '');
  return { integer, fraction: text.slice(text.length - decimals) };
}

/** .NET "G15"-style scientific text: `1.5E-11`, `1.23456789012346E+15`. */
export function scientificG15(value: number): string {
  const { digits, exponent } = significantDigits(Math.abs(value));
  const mantissaTail = digits.slice(1).replace(/0+$/, '');
  const mantissa = digits[0]! + (mantissaTail ? '.' + mantissaTail : '');
  const exp = `${exponent < 0 ? '-' : '+'}${String(Math.abs(exponent)).padStart(2, '0')}`;
  return `${value < 0 ? '-' : ''}${mantissa}E${exp}`;
}
