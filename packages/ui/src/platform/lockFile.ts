/**
 * LibreOffice's lock files. While LibreOffice has `foo.ods` open it keeps `.~lock.foo.ods#` beside
 * it: one record of five comma-separated fields closed with `;` (user name, system user, host,
 * time, user URL), with `\`, `,` and `;` escaped by a backslash inside a field. Cellestra only
 * reads them, to warn on open; it never writes or deletes one.
 */

/** Who has the file open in LibreOffice, as far as the lock file says. */
export interface LockOwner {
  /** The LibreOffice user name, else the system user or the host; null when the lock names no one. */
  readonly name: string | null;
  /** When they opened it, as LibreOffice wrote it (`14.11.2026 10:22`); null when missing. */
  readonly date: string | null;
}

/** The owner a lock file's text names. Never throws: an unreadable lock still names an owner (no one). */
export function parseLockFile(text: string): LockOwner {
  const fields = firstRecord(text.replace(/^\uFEFF/, '')).map((f) => f.trim());
  const at = (i: number) => (fields[i] ? fields[i] : null);
  return { name: at(0) ?? at(1) ?? at(2), date: at(3) };
}

/** The warning shown when a document opens with a lock beside it. */
export function lockWarning(owner: LockOwner): string {
  const who = owner.name ? `Opened by ${owner.name}` : 'Open';
  const when = owner.date ? ` on ${owner.date}` : '';
  return `${who}${when} in LibreOffice – your changes may conflict.`;
}

/** The fields of the first `;`-terminated record, unescaped. */
function firstRecord(text: string): string[] {
  const fields: string[] = [];
  let field = '';
  for (let i = 0; i < text.length; i++) {
    const c = text[i]!;
    if (c === '\\' && i + 1 < text.length) field += text[++i];
    else if (c === ',') {
      fields.push(field);
      field = '';
    } else if (c === ';') break;
    else field += c;
  }
  fields.push(field);
  return fields;
}
