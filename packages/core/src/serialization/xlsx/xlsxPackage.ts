import { OdsPackage, OdsPackageTooLargeError, OdsUnsupportedFileError } from '../ods/odsPackage.ts';
import { tryParseXml, type XmlAttribute, type XmlElement } from '../../xml/index.ts';

/**
 * The OLE compound-file signature. A password-protected `.xlsx` is an encrypted OLE file, and so is
 * an old binary `.xls`; neither is a ZIP package.
 */
const OLE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** Relationship types end in these names, in both the transitional and the strict namespace. */
export const RelType = {
  officeDocument: '/officeDocument',
  worksheet: '/worksheet',
  chartsheet: '/chartsheet',
  dialogsheet: '/dialogsheet',
  macrosheet: '/xlMacrosheet',
  sharedStrings: '/sharedStrings',
  styles: '/styles',
  theme: '/theme',
  hyperlink: '/hyperlink',
  comments: '/comments',
  pivotTable: '/pivotTable',
  vbaProject: '/vbaProject',
} as const;

/** One relationship from a part's `_rels` file. */
export interface Relationship {
  readonly type: string;
  /** The package path for an internal target, or the URL as written for an external one. */
  readonly target: string;
  readonly external: boolean;
}

/** An element's children with the given local name, whatever their namespace (transitional or strict). */
export function kids(el: XmlElement | null, local: string): XmlElement[] {
  return el ? el.children().filter((c) => c.local === local) : [];
}

export function kid(el: XmlElement | null, local: string): XmlElement | null {
  return el?.children().find((c) => c.local === local) ?? null;
}

/** An unprefixed attribute's value. */
export function at(el: XmlElement | null, local: string): string | null {
  return el?.attr('', local) ?? null;
}

/** An unprefixed attribute from a streamed element's attribute list. */
export function attrIn(attributes: readonly XmlAttribute[], local: string): string | null {
  for (const a of attributes) if (a.name.local === local && a.name.uri === '') return a.value;
  return null;
}

/** The `r:id` of an element (an attribute named `id` in the relationships namespace). */
export function relIdIn(attributes: readonly XmlAttribute[]): string | null {
  for (const a of attributes)
    if (a.name.local === 'id' && a.name.uri.endsWith('/relationships')) return a.value;
  return null;
}

/** An XML boolean as OOXML writes it: absent means `fallback`, else `1`/`true` are true. */
export function flag(value: string | null, fallback = false): boolean {
  if (value === null) return fallback;
  return value === '1' || value.toLowerCase() === 'true';
}

/** Thrown for bytes that aren't an `.xlsx`; the message is for the person opening the file. */
function unsupported(message: string): OdsUnsupportedFileError {
  return new OdsUnsupportedFileError(message);
}

/** The ZIP package of an `.xlsx`, with its relationships resolved to part paths. */
export class XlsxPackage {
  readonly #pkg: OdsPackage;

  private constructor(pkg: OdsPackage) {
    this.#pkg = pkg;
  }

  /**
   * Open the package. Refuses an OLE file (an encrypted workbook or a binary `.xls`) and bytes that
   * aren't a ZIP. Worksheets and shared strings stay compressed and are inflated as they are read.
   */
  static read(bytes: Uint8Array): XlsxPackage {
    if (OLE_SIGNATURE.every((b, i) => bytes[i] === b)) {
      throw unsupported(
        'This file is password-protected or in the old .xls format. Save it as .xlsx without a password, then open it here.',
      );
    }
    const streamed = (name: string): boolean =>
      /^xl\/worksheets\/[^/]+\.xml$/i.test(name) || /sharedStrings\.xml$/i.test(name);
    try {
      return new XlsxPackage(OdsPackage.read(bytes, undefined, streamed));
    } catch (e) {
      if (e instanceof OdsPackageTooLargeError) throw e;
      throw unsupported('This file is damaged or is not an .xlsx workbook.');
    }
  }

  has(path: string): boolean {
    return this.#pkg.contains(path);
  }

  /** A whole part parsed as XML, or null when it is missing. Throws on malformed XML. */
  xml(path: string): XmlElement | null {
    const text = this.#pkg.getText(path);
    if (text === null) return null;
    const doc = tryParseXml(text, { preserveWhitespace: true });
    if (!doc) throw unsupported(`This file is damaged: ${path} is not valid XML.`);
    return doc.root;
  }

  /** A part's text in pieces, inflated as they are taken; null when missing. */
  chunks(path: string): Iterable<string> | null {
    return this.#pkg.textChunks(path);
  }

  /** The relationships of `part`, by id. */
  rels(part: string): Map<string, Relationship> {
    const slash = part.lastIndexOf('/');
    const dir = part.slice(0, slash + 1);
    const relsPath = `${dir}_rels/${part.slice(slash + 1)}.rels`;
    const out = new Map<string, Relationship>();
    for (const rel of kids(this.xml(relsPath), 'Relationship')) {
      const id = at(rel, 'Id');
      const target = at(rel, 'Target');
      if (id === null || target === null) continue;
      const external = at(rel, 'TargetMode') === 'External';
      out.set(id, {
        type: at(rel, 'Type') ?? '',
        target: external ? target : resolvePath(dir, target),
        external,
      });
    }
    return out;
  }

  /** The workbook part's path, from the package relationships. */
  workbookPath(): string {
    for (const rel of this.rels('').values())
      if (rel.type.endsWith(RelType.officeDocument)) return rel.target;
    return 'xl/workbook.xml';
  }
}

/** Resolve a relationship target against the folder of the part that names it. */
export function resolvePath(dir: string, target: string): string {
  const parts = (target.startsWith('/') ? target.slice(1) : dir + target).split('/');
  const out: string[] = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}
