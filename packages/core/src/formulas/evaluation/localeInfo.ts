/**
 * Locale-specific symbols and names read through `Intl` (no host globals), cached per locale.
 * Shared by the number formatter, VALUE/DATEVALUE-style parsing, and criteria matching.
 */
export interface LocaleInfo {
  readonly locale: string;
  readonly decimal: string;
  readonly group: string;
  readonly currencySymbol: string;
  /** Full month names, January first. */
  readonly months: readonly string[];
  readonly monthsShort: readonly string[];
  /** Full day names, Sunday first. */
  readonly days: readonly string[];
  readonly daysShort: readonly string[];
  readonly am: string;
  readonly pm: string;
  /** Order of the numeric short-date fields, e.g. ['month', 'day', 'year'] for en-US. */
  readonly dateOrder: readonly ('day' | 'month' | 'year')[];
}

const cache = new Map<string, LocaleInfo>();

/** A locale tag Intl accepts: the given one when valid, otherwise 'en-US'. */
export function safeLocale(locale: string): string {
  try {
    return Intl.getCanonicalLocales(locale)[0] ?? 'en-US';
  } catch {
    return 'en-US';
  }
}

/** Region → currency for the common cases (Intl offers no locale → currency mapping). */
const REGION_CURRENCY: Readonly<Record<string, string>> = {
  US: 'USD',
  CA: 'CAD',
  AU: 'AUD',
  NZ: 'NZD',
  GB: 'GBP',
  IE: 'EUR',
  DE: 'EUR',
  AT: 'EUR',
  FR: 'EUR',
  ES: 'EUR',
  IT: 'EUR',
  NL: 'EUR',
  BE: 'EUR',
  PT: 'EUR',
  FI: 'EUR',
  GR: 'EUR',
  LU: 'EUR',
  CH: 'CHF',
  JP: 'JPY',
  CN: 'CNY',
  IN: 'INR',
  BR: 'BRL',
  MX: 'MXN',
  SE: 'SEK',
  NO: 'NOK',
  DK: 'DKK',
  PL: 'PLN',
  CZ: 'CZK',
  RU: 'RUB',
  KR: 'KRW',
  ZA: 'ZAR',
  TR: 'TRY',
};

function currencySymbolFor(locale: string): string {
  let region: string | undefined;
  try {
    region = new Intl.Locale(locale).maximize().region;
  } catch {
    region = undefined;
  }
  const currency = region ? REGION_CURRENCY[region] : undefined;
  if (!currency) return '¤';
  const part = new Intl.NumberFormat(locale, { style: 'currency', currency })
    .formatToParts(1)
    .find((p) => p.type === 'currency');
  return part?.value ?? '¤';
}

function utc(year: number, month: number, day: number, hour = 0): Date {
  return new Date(Date.UTC(year, month - 1, day, hour));
}

function names(
  locale: string,
  options: Intl.DateTimeFormatOptions,
  count: number,
  at: (i: number) => Date,
): string[] {
  const fmt = new Intl.DateTimeFormat(locale, { ...options, timeZone: 'UTC' });
  return Array.from({ length: count }, (_, i) => fmt.format(at(i)));
}

function build(locale: string): LocaleInfo {
  const parts = new Intl.NumberFormat(locale, { useGrouping: true }).formatToParts(1234567.5);
  const decimal = parts.find((p) => p.type === 'decimal')?.value ?? '.';
  const group = parts.find((p) => p.type === 'group')?.value ?? ',';
  const monthAt = (i: number): Date => utc(2001, i + 1, 15);
  const dayAt = (i: number): Date => utc(2001, 1, 7 + i); // 2001-01-07 was a Sunday
  const dayPeriod = (hour: number): string =>
    new Intl.DateTimeFormat(locale, { hour: 'numeric', hour12: true, timeZone: 'UTC' })
      .formatToParts(utc(2001, 1, 1, hour))
      .find((p) => p.type === 'dayPeriod')?.value ?? (hour < 12 ? 'AM' : 'PM');
  const order: ('day' | 'month' | 'year')[] = new Intl.DateTimeFormat(locale, {
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    timeZone: 'UTC',
  })
    .formatToParts(utc(2001, 2, 3))
    .flatMap((p) => (p.type === 'day' || p.type === 'month' || p.type === 'year' ? [p.type] : []));
  return Object.freeze({
    locale,
    decimal,
    group,
    currencySymbol: currencySymbolFor(locale),
    months: names(locale, { month: 'long' }, 12, monthAt),
    monthsShort: names(locale, { month: 'short' }, 12, monthAt),
    days: names(locale, { weekday: 'long' }, 7, dayAt),
    daysShort: names(locale, { weekday: 'short' }, 7, dayAt),
    am: dayPeriod(9),
    pm: dayPeriod(15),
    dateOrder: order.length === 3 ? order : (['month', 'day', 'year'] as const),
  });
}

/** Cached `LocaleInfo` for a BCP-47 tag (invalid tags fall back to en-US). */
export function localeInfo(locale: string): LocaleInfo {
  let info = cache.get(locale);
  if (!info) {
    info = build(safeLocale(locale));
    cache.set(locale, info);
  }
  return info;
}
