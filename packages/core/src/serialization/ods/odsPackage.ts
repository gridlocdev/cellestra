import { DecodeUTF8, Inflate, strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate';
import { GeneratedParts, OdsXmlNames } from './odsXmlNames.ts';

/** The most a package may expand to when read (the sum of its entries' uncompressed sizes). */
export const MAX_UNCOMPRESSED_BYTES = 512 * 2 ** 20;

/** How much compressed data a streamed part is inflated from at a time. */
const STREAM_CHUNK_BYTES = 64 * 1024;

/**
 * Parts kept compressed when read and inflated only as they are streamed: content.xml holds the
 * cells, and inflated it can be a hundred times the size of the file.
 */
const STREAMED_PARTS: ReadonlySet<string> = new Set([OdsXmlNames.contentPart]);

/** A part still as it was in the ZIP file: `method` 0 is stored, 8 is deflated. */
interface CompressedPart {
  readonly data: Uint8Array;
  readonly method: number;
  /** The most it may inflate to: what the package cap leaves after every other part. */
  readonly limit: number;
  /** The package cap, for the message when the part goes past `limit`. */
  readonly cap: number;
}

/** Thrown when a package's entries would inflate past `MAX_UNCOMPRESSED_BYTES`. */
export class OdsPackageTooLargeError extends Error {
  override readonly name = 'OdsPackageTooLargeError';
}

/**
 * Thrown when the bytes are not a spreadsheet Cellestra can open: not a ZIP package, another kind
 * of document, or password-protected. The message is written for the person opening the file.
 */
export class OdsUnsupportedFileError extends Error {
  override readonly name = 'OdsUnsupportedFileError';
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
  private readonly parts = new Map<string, Uint8Array | CompressedPart>();

  /** Unmodeled `<office:spreadsheet>` children captured on read, re-emitted on write (PRD §8.1). */
  readonly foreignSpreadsheetChildren: ForeignXml[] = [];

  get partNames(): string[] {
    return [...this.parts.keys()];
  }

  contains(name: string): boolean {
    return this.parts.has(name);
  }

  getBytes(name: string): Uint8Array | null {
    const part = this.parts.get(name);
    if (!part || part instanceof Uint8Array) return part ?? null;
    return inflatePart(part);
  }

  /** The part decoded as UTF-8 (a leading BOM is dropped), or null when absent. */
  getText(name: string): string | null {
    const bytes = this.getBytes(name);
    if (!bytes) return null;
    const text = strFromU8(bytes);
    return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  }

  /**
   * The part decoded as UTF-8 in pieces (a leading BOM is dropped), or null when absent. A part
   * read from a file is inflated a little at a time as the pieces are taken, so the whole of it is
   * never in memory at once. Throws `OdsUnsupportedFileError` when the data is damaged and
   * `OdsPackageTooLargeError` when it inflates past the package cap.
   */
  textChunks(name: string): Iterable<string> | null {
    const part = this.parts.get(name);
    if (!part) return null;
    if (part instanceof Uint8Array) return [this.getText(name)!];
    return streamText(part);
  }

  setText(name: string, content: string): void {
    this.parts.set(name, strToU8(content));
  }

  setBytes(name: string, content: Uint8Array): void {
    this.parts.set(name, content);
  }

  /** Drop a part; true when it was there. */
  remove(name: string): boolean {
    return this.parts.delete(name);
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
    let entries: Record<string, Uint8Array>;
    const kept = new Map<string, number>();
    try {
      entries = unzipSync(bytes, {
        filter: (file) => {
          total += file.originalSize;
          if (total > maxUncompressedBytes) {
            throw new OdsPackageTooLargeError(
              `The file expands to more than ${Math.round(maxUncompressedBytes / 2 ** 20)} MiB.`,
            );
          }
          // A streamed part stays compressed; it is found in the ZIP directory below.
          if (STREAMED_PARTS.has(file.name) && (file.compression === 0 || file.compression === 8)) {
            kept.set(file.name, file.originalSize);
            return false;
          }
          return true;
        },
      });
    } catch (e) {
      if (e instanceof OdsPackageTooLargeError) throw e;
      throw new OdsUnsupportedFileError('This file is damaged or is not an .ods spreadsheet.');
    }
    for (const [name, data] of Object.entries(entries)) {
      if (name.endsWith('/')) continue; // directory entry
      pkg.parts.set(name, data);
    }
    for (const [name, size] of kept) {
      const found = findCompressedEntry(bytes, name);
      // A layout this reader does not know (ZIP64, say) is read whole, as every other part is.
      if (found) {
        const limit = maxUncompressedBytes - (total - size);
        // A copy, so the file's bytes need not be kept for as long as the package is.
        pkg.parts.set(name, {
          data: found.data.slice(),
          method: found.method,
          limit,
          cap: maxUncompressedBytes,
        });
        continue;
      }
      try {
        const whole = unzipSync(bytes, { filter: (file) => file.name === name })[name];
        if (whole) pkg.parts.set(name, whole);
      } catch {
        throw new OdsUnsupportedFileError('This file is damaged or is not an .ods spreadsheet.');
      }
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
    for (const name of this.parts.keys()) {
      if (name === OdsXmlNames.mimetypePart) continue;
      files[name] = [this.getBytes(name)!, { level: 6 }];
    }
    return zipSync(files);
  }
}

function damaged(): OdsUnsupportedFileError {
  return new OdsUnsupportedFileError('This file is damaged or is not an .ods spreadsheet.');
}

function inflatePart(part: CompressedPart): Uint8Array {
  const chunks = [...inflatedPieces(part)];
  if (chunks.length === 1) return chunks[0]!;
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
  let at = 0;
  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.length;
  }
  return out;
}

/**
 * Inflate `part` a piece of input at a time, yielding the output as it comes, so taking the
 * pieces one by one never holds the whole part. The running size is checked against the limit.
 */
function* inflatedPieces(part: CompressedPart): Generator<Uint8Array> {
  const { data } = part;
  const out: Uint8Array[] = [];
  const inflater = part.method === 0 ? null : new Inflate((chunk) => out.push(chunk));
  let size = 0;
  for (let at = 0; at < data.length || at === 0; at += STREAM_CHUNK_BYTES) {
    const piece = data.subarray(at, at + STREAM_CHUNK_BYTES);
    if (!inflater) out.push(piece);
    else {
      try {
        inflater.push(piece, at + STREAM_CHUNK_BYTES >= data.length);
      } catch {
        throw damaged();
      }
    }
    for (const chunk of out) {
      size += chunk.length;
      if (size > part.limit) {
        throw new OdsPackageTooLargeError(
          `The file expands to more than ${Math.round(part.cap / 2 ** 20)} MiB.`,
        );
      }
    }
    yield* out;
    out.length = 0;
  }
}

function* streamText(part: CompressedPart): Generator<string> {
  const pending: string[] = [];
  let first = true;
  const decoder = new DecodeUTF8((text) => {
    if (first && text.length > 0) {
      first = false;
      if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    }
    if (text.length > 0) pending.push(text);
  });
  for (const chunk of inflatedPieces(part)) {
    decoder.push(chunk);
    yield* pending;
    pending.length = 0;
  }
  decoder.push(new Uint8Array(0), true);
  yield* pending;
}

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const LOCAL_SIGNATURE = 0x04034b50;
/** End of central directory record size, and the longest comment that may follow it. */
const EOCD_SIZE = 22;
const MAX_COMMENT = 0xffff;
/** A 32-bit size or offset with every bit set means the real value is in a ZIP64 record. */
const ZIP64_MARKER = 0xffffffff;

/**
 * The compressed data of the last entry named `name`, found through the ZIP central directory
 * (whose sizes are right even when the local headers defer them to a data descriptor). Null when
 * the entry is missing or the archive uses a layout this does not read (ZIP64).
 */
function findCompressedEntry(zip: Uint8Array, name: string): { data: Uint8Array; method: number } | null {
  const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  const u16 = (at: number): number => view.getUint16(at, true);
  const u32 = (at: number): number => view.getUint32(at, true);
  const fits = (at: number, length: number): boolean => at >= 0 && at + length <= zip.length;

  let eocd = -1;
  for (let at = zip.length - EOCD_SIZE; at >= Math.max(0, zip.length - EOCD_SIZE - MAX_COMMENT); at--) {
    if (u32(at) === EOCD_SIGNATURE) {
      eocd = at;
      break;
    }
  }
  if (eocd < 0) return null;
  const count = u16(eocd + 10);
  let at = u32(eocd + 16);
  if (at === ZIP64_MARKER) return null;

  const wanted = strToU8(name);
  let found: { data: Uint8Array; method: number } | null = null;
  for (let i = 0; i < count; i++) {
    if (!fits(at, 46) || u32(at) !== CENTRAL_SIGNATURE) return null;
    const method = u16(at + 10);
    const compressedSize = u32(at + 20);
    const nameLength = u16(at + 28);
    const next = at + 46 + nameLength + u16(at + 30) + u16(at + 32);
    const local = u32(at + 42);
    if (!fits(at + 46, nameLength)) return null;
    const entryName = zip.subarray(at + 46, at + 46 + nameLength);
    if (entryName.length === wanted.length && entryName.every((b, k) => b === wanted[k])) {
      if (compressedSize === ZIP64_MARKER || local === ZIP64_MARKER) return null;
      if (!fits(local, 30) || u32(local) !== LOCAL_SIGNATURE) return null;
      const start = local + 30 + u16(local + 26) + u16(local + 28);
      if (!fits(start, compressedSize)) return null;
      found = { data: zip.subarray(start, start + compressedSize), method };
    }
    at = next;
  }
  return found;
}
