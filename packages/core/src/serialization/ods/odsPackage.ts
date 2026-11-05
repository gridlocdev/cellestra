import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate';
import { GeneratedParts, OdsXmlNames } from './odsXmlNames.ts';

/** The most a package may expand to when read (the sum of its entries' uncompressed sizes). */
export const MAX_UNCOMPRESSED_BYTES = 512 * 2 ** 20;

/** Thrown when a package's entries would inflate past `MAX_UNCOMPRESSED_BYTES`. */
export class OdsPackageTooLargeError extends Error {
  override readonly name = 'OdsPackageTooLargeError';
}

/** An unmodeled child of `<office:spreadsheet>`, kept verbatim (PRD §8.1). */
export interface ForeignXml {
  readonly beforeTables: boolean;
  /** The element serialized standalone — every namespace it uses is declared on it. */
  readonly xml: string;
}

/**
 * The `.ods` ZIP container. Enforces the ODF rule that `mimetype` is the **first** entry and
 * **stored uncompressed** (Golden Rule #7). Holds every part as bytes so parts Cellestra doesn't
 * model are preserved verbatim on save (PRD §8.1).
 */
export class OdsPackage {
  private readonly parts = new Map<string, Uint8Array>();

  /** Unmodeled `<office:spreadsheet>` children captured on read, re-emitted on write (PRD §8.1). */
  readonly foreignSpreadsheetChildren: ForeignXml[] = [];

  get partNames(): string[] {
    return [...this.parts.keys()];
  }

  contains(name: string): boolean {
    return this.parts.has(name);
  }

  getBytes(name: string): Uint8Array | null {
    return this.parts.get(name) ?? null;
  }

  /** The part decoded as UTF-8 (a leading BOM is dropped), or null when absent. */
  getText(name: string): string | null {
    const bytes = this.parts.get(name);
    if (!bytes) return null;
    const text = strFromU8(bytes);
    return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  }

  setText(name: string, content: string): void {
    this.parts.set(name, strToU8(content));
  }

  setBytes(name: string, content: Uint8Array): void {
    this.parts.set(name, content);
  }

  /** The parts that are NOT regenerated on save — preserved as-is (charts, images, etc.). */
  get preservedParts(): string[] {
    return this.partNames.filter((name) => !GeneratedParts.has(name));
  }

  static read(bytes: Uint8Array, maxUncompressedBytes = MAX_UNCOMPRESSED_BYTES): OdsPackage {
    const pkg = new OdsPackage();
    // A zip bomb declares tiny compressed entries that inflate to gigabytes. fflate inflates each
    // entry into a buffer of its declared size (and never grows it), so bounding the declared sizes
    // bounds the memory reading can take.
    let total = 0;
    const entries = unzipSync(bytes, {
      filter: (file) => {
        total += file.originalSize;
        if (total > maxUncompressedBytes) {
          throw new OdsPackageTooLargeError(
            `The file expands to more than ${Math.round(maxUncompressedBytes / 2 ** 20)} MiB.`,
          );
        }
        return true;
      },
    });
    for (const [name, data] of Object.entries(entries)) {
      if (name.endsWith('/')) continue; // directory entry
      pkg.parts.set(name, data);
    }
    return pkg;
  }

  write(): Uint8Array {
    // The mimetype MUST be first and stored (uncompressed), with the exact bytes. zipSync writes
    // entries in key order; string keys keep insertion order (a part named like an array index —
    // "0", "17" — would sort first, and no ODF part is ever named that way).
    const files: Zippable = {
      [OdsXmlNames.mimetypePart]: [strToU8(OdsXmlNames.mimeType), { level: 0 }],
    };
    for (const [name, data] of this.parts) {
      if (name === OdsXmlNames.mimetypePart) continue;
      files[name] = [data, { level: 6 }];
    }
    return zipSync(files);
  }
}
