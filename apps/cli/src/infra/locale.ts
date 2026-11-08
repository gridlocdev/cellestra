import { CliError } from './errors.ts';

/**
 * `--locale` → a BCP 47 tag the engine understands, or null to fall back to the document locale
 * (else the system's). Unknown tags are rejected up front rather than silently formatting as en-US.
 */
export function resolveLocale(name: string | null | undefined): string | null {
  if (name === null || name === undefined || name.trim() === '') return null;
  let canonical: string | undefined;
  try {
    canonical = Intl.getCanonicalLocales(name.trim())[0];
  } catch {
    canonical = undefined;
  }
  if (!canonical || Intl.NumberFormat.supportedLocalesOf([canonical]).length === 0) {
    throw new CliError(`Unknown locale '${name}' (try a tag like en-US or de-DE).`);
  }
  return canonical;
}
