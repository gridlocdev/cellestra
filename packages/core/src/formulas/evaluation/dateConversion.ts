/**
 * A calendar date and wall-clock time with no time zone (the .NET `DateTimeKind.Unspecified` analogue).
 * `month` is 1–12. Proleptic Gregorian calendar.
 */
export interface CivilDateTime {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly hour: number;
  readonly minute: number;
  readonly second: number;
  readonly millisecond: number;
}

const MS_PER_DAY = 86_400_000;

/** Days since 1970-01-01 for a proleptic-Gregorian civil date (Howard Hinnant's algorithm). */
export function daysFromCivil(year: number, month: number, day: number): number {
  const y = month <= 2 ? year - 1 : year;
  const era = Math.floor(y / 400);
  const yoe = y - era * 400;
  const mp = (month + 9) % 12;
  const doy = Math.floor((153 * mp + 2) / 5) + day - 1;
  const doe = yoe * 365 + Math.floor(yoe / 4) - Math.floor(yoe / 100) + doy;
  return era * 146097 + doe - 719468;
}

/** Inverse of `daysFromCivil`. */
export function civilFromDays(days: number): { year: number; month: number; day: number } {
  const z = days + 719468;
  const era = Math.floor(z / 146097);
  const doe = z - era * 146097;
  const yoe = Math.floor(
    (doe - Math.floor(doe / 1460) + Math.floor(doe / 36524) - Math.floor(doe / 146096)) / 365,
  );
  const doy = doe - (365 * yoe + Math.floor(yoe / 4) - Math.floor(yoe / 100));
  const mp = Math.floor((5 * doy + 2) / 153);
  const day = doy - Math.floor((153 * mp + 2) / 5) + 1;
  const month = mp < 10 ? mp + 3 : mp - 9;
  return { year: yoe + era * 400 + (month <= 2 ? 1 : 0), month, day };
}

/** Day of week for a day count since 1970-01-01: Sunday = 0 … Saturday = 6. */
export function dayOfWeekFromDays(days: number): number {
  return (((days + 4) % 7) + 7) % 7;
}

const EPOCH_DAYS = daysFromCivil(1899, 12, 30);
/** 0001-01-01 and the end of 9999-12-31, as serials (the .NET DateTime span). */
const MIN_SERIAL = daysFromCivil(1, 1, 1) - EPOCH_DAYS;
const MAX_SERIAL_EXCLUSIVE = daysFromCivil(10000, 1, 1) - EPOCH_DAYS;

function fromTotalMs(totalMs: number): CivilDateTime {
  const days = Math.floor(totalMs / MS_PER_DAY);
  let rest = totalMs - days * MS_PER_DAY;
  const { year, month, day } = civilFromDays(days + EPOCH_DAYS);
  const hour = Math.floor(rest / 3_600_000);
  rest -= hour * 3_600_000;
  const minute = Math.floor(rest / 60_000);
  rest -= minute * 60_000;
  const second = Math.floor(rest / 1000);
  return { year, month, day, hour, minute, second, millisecond: rest - second * 1000 };
}

function validateYear(year: number): void {
  if (!Number.isInteger(year) || year < 1 || year > 9999) throw new RangeError(`Year out of range: ${year}`);
}

/**
 * Date/time serial numbers on the ODF/Calc epoch **1899-12-30** (serial 0): whole days are integers,
 * the time of day is the fraction. Pure arithmetic — the host time zone never matters.
 */
export const DateConversion = {
  epoch: Object.freeze({
    year: 1899,
    month: 12,
    day: 30,
    hour: 0,
    minute: 0,
    second: 0,
    millisecond: 0,
  }) as CivilDateTime,

  toSerial(d: CivilDateTime): number {
    const days = daysFromCivil(d.year, d.month, d.day) - EPOCH_DAYS;
    const ms = ((d.hour * 60 + d.minute) * 60 + d.second) * 1000 + d.millisecond;
    return (days * MS_PER_DAY + ms) / MS_PER_DAY;
  },

  /**
   * DATE(year, month, day). A two-digit year (0–99) resolves on Calc's 1930–2029 window; 100–1899 is
   * literal. Months and days overflow/underflow and roll over. Throws RangeError outside years 1–9999.
   */
  dateToSerial(year: number, month: number, day: number): number {
    if (year >= 0 && year <= 99) year += year <= 29 ? 2000 : 1900;
    validateYear(year);
    const totalMonths = year * 12 + (month - 1);
    const y = Math.floor(totalMonths / 12);
    validateYear(y);
    const m = totalMonths - y * 12 + 1;
    const serial = daysFromCivil(y, m, 1) + (day - 1) - EPOCH_DAYS;
    if (serial < MIN_SERIAL || serial >= MAX_SERIAL_EXCLUSIVE) throw new RangeError('Date out of range.');
    return serial;
  },

  /** Serial → date and time, rounded to the nearest millisecond. */
  fromSerial(serial: number): CivilDateTime {
    return fromTotalMs(Math.round(serial * MS_PER_DAY));
  },

  /** The whole-day part of a serial (floored), at midnight. */
  dateFromSerial(serial: number): CivilDateTime {
    return fromTotalMs(Math.floor(serial) * MS_PER_DAY);
  },

  /**
   * Convert without throwing: null for NaN/infinite serials or ones outside years 1–9999, so a date
   * format can fall back to a plain number. `dateOnly` drops the time-of-day fraction.
   */
  tryToDateTime(serial: number, dateOnly: boolean): CivilDateTime | null {
    if (!Number.isFinite(serial)) return null;
    const days = dateOnly ? Math.floor(serial) : serial;
    if (days < MIN_SERIAL || days >= MAX_SERIAL_EXCLUSIVE) return null;
    const totalMs = Math.round(days * MS_PER_DAY);
    if (totalMs >= MAX_SERIAL_EXCLUSIVE * MS_PER_DAY) return null;
    return fromTotalMs(totalMs);
  },

  /** Day of week of a serial's date: Sunday = 0 … Saturday = 6. */
  dayOfWeek(serial: number): number {
    return dayOfWeekFromDays(Math.floor(serial) + EPOCH_DAYS);
  },
} as const;

/**
 * The host's current local wall-clock time, read through `Intl` (no local-time `Date` getters). Used as
 * the default clock for TODAY/NOW; tests inject their own.
 */
export function systemClock(): CivilDateTime {
  const parts = clockFormat().formatToParts(new Date());
  const get = (t: Intl.DateTimeFormatPartTypes): number =>
    Number(parts.find((p) => p.type === t)?.value ?? 0);
  return {
    year: get('year'),
    month: get('month'),
    day: get('day'),
    hour: get('hour') % 24,
    minute: get('minute'),
    second: get('second'),
    millisecond: 0,
  };
}

let clockFormatter: Intl.DateTimeFormat | null = null;
function clockFormat(): Intl.DateTimeFormat {
  clockFormatter ??= new Intl.DateTimeFormat('en-US', {
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    hour: 'numeric',
    minute: 'numeric',
    second: 'numeric',
    hourCycle: 'h23',
  });
  return clockFormatter;
}
