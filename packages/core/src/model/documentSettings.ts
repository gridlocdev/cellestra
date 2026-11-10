/** How formulas recalculate (PRD §8.6). */
export type CalculationMode = 'automatic' | 'manual';

/** Fallback when neither the document nor the host names a locale. */
let hostLocale = 'en-US';

/** Hosts (the UI) call this once with `navigator.language`; core never reads browser globals. */
export function setHostLocale(locale: string): void {
  hostLocale = locale;
}

export function getHostLocale(): string {
  return hostLocale;
}

/**
 * Document-wide settings persisted in `.ods` `settings.xml`: locale (PRD §8.3 — document locale
 * when present, else the system's), calculation mode, active sheet, zoom.
 */
export class DocumentSettings {
  /** BCP 47 tag, or null = unspecified (fall back to the host locale). */
  locale: string | null = null;
  calculationMode: CalculationMode = 'automatic';
  activeSheetIndex = 0;
  /** The view zoom in percent (Calc's `ZoomValue`), or null when the file had none. */
  zoom: number | null = null;

  get effectiveLocale(): string {
    return this.locale ?? hostLocale;
  }
}
