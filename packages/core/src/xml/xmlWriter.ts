import type { XmlDocument } from './xmlTree.ts';
import {
  XML_NS_URI,
  XMLNS_URI,
  XmlComment,
  XmlElement,
  XmlProcessingInstruction,
  XmlText,
  type XmlName,
  type XmlNode,
} from './xmlTree.ts';

export interface XmlWriteOptions {
  /**
   * Pretty-print element-only content two spaces per level (C#'s `XmlWriterSettings.Indent`).
   * Mixed content — an element with any text child — is written exactly as stored, so whitespace
   * that is data (a `text:p`) is never changed. Default true.
   */
  readonly indent?: boolean;
  /** Write the `<?xml …?>` declaration (documents only). Default true. */
  readonly declaration?: boolean;
}

/** Serialize a document. Namespace declarations the tree needs but lacks are added where needed. */
export function serializeXml(doc: XmlDocument, options: XmlWriteOptions = {}): string {
  const indent = options.indent ?? true;
  let out = '';
  if (options.declaration ?? true) {
    const d = doc.declaration ?? { version: '1.0', encoding: 'UTF-8' };
    out += `<?xml version="${d.version}"`;
    out += ` encoding="${d.encoding ?? 'UTF-8'}"`;
    if (d.standalone) out += ` standalone="${d.standalone}"`;
    out += '?>';
    if (indent) out += '\n';
  }
  for (const node of doc.prolog) {
    out += writeLeaf(node);
    if (indent) out += '\n';
  }
  const w = new Writer(indent);
  w.element(doc.root, new Map(), 0);
  return out + w.out;
}

/**
 * Serialize one element standalone — every namespace it uses is declared on it, so the string parses
 * on its own (C#'s `XElement.ToString()`). Not indented by default: what was read is what is written.
 */
export function serializeXmlElement(element: XmlElement, options: { indent?: boolean } = {}): string {
  const w = new Writer(options.indent ?? false);
  w.element(element, new Map(), 0);
  return w.out;
}

interface Frame {
  readonly e: XmlElement;
  readonly scope: Map<string, string>;
  readonly depth: number;
  readonly name: string;
  readonly indentChildren: boolean;
  i: number;
}

/** Writes elements with an explicit stack, so nesting depth cannot overflow the call stack. */
class Writer {
  out = '';
  private readonly indent: boolean;
  constructor(indent: boolean) {
    this.indent = indent;
  }

  element(root: XmlElement, inherited: Map<string, string>, depth: number, pretty = this.indent): void {
    const stack: Frame[] = [];
    const open = (e: XmlElement, scopeIn: Map<string, string>, d: number, p: boolean): void => {
      const frame = this.openTag(e, scopeIn, d, p);
      if (frame) stack.push(frame);
    };
    open(root, inherited, depth, pretty);
    while (stack.length > 0) {
      const top = stack[stack.length - 1]!;
      if (top.i >= top.e.nodes.length) {
        if (top.indentChildren) this.out += '\n' + '  '.repeat(top.depth);
        this.out += `</${top.name}>`;
        stack.pop();
        continue;
      }
      const n = top.e.nodes[top.i++]!;
      if (top.indentChildren) this.out += '\n' + '  '.repeat(top.depth + 1);
      if (n instanceof XmlElement) open(n, top.scope, top.depth + 1, top.indentChildren);
      else this.out += writeLeaf(n);
    }
  }

  /** Write the start tag (or the whole empty element); returns the frame for its children. */
  private openTag(
    e: XmlElement,
    inherited: Map<string, string>,
    depth: number,
    pretty: boolean,
  ): Frame | null {
    const scope = new Map(inherited);
    // Declarations the element carries itself come into scope first.
    for (const a of e.attributes) {
      if (a.name.uri !== XMLNS_URI) continue;
      scope.set(a.name.prefix === 'xmlns' ? a.name.local : '', a.value);
    }

    const extra: string[] = [];
    const declare = (prefix: string, uri: string): void => {
      scope.set(prefix, uri);
      extra.push(prefix ? ` xmlns:${prefix}="${escapeAttr(uri)}"` : ` xmlns="${escapeAttr(uri)}"`);
    };

    const elementName = resolveElement(e.name, scope, declare);
    let attrs = '';
    for (const a of e.attributes) {
      if (a.name.uri === XMLNS_URI) {
        const q = a.name.prefix === 'xmlns' ? `xmlns:${a.name.local}` : 'xmlns';
        attrs += ` ${q}="${escapeAttr(a.value)}"`;
      }
    }
    let regular = '';
    for (const a of e.attributes) {
      if (a.name.uri === XMLNS_URI) continue;
      regular += ` ${resolveAttribute(a.name, scope, declare)}="${escapeAttr(a.value)}"`;
    }

    this.out += `<${elementName}${attrs}${extra.join('')}${regular}`;
    if (e.nodes.length === 0) {
      this.out += '/>';
      return null;
    }
    this.out += '>';

    const mixed = e.nodes.some((n) => n instanceof XmlText);
    return { e, scope, depth, name: elementName, indentChildren: pretty && !mixed, i: 0 };
  }
}

type Declare = (prefix: string, uri: string) => void;

function resolveElement(name: XmlName, scope: Map<string, string>, declare: Declare): string {
  if (name.uri === '') {
    if (name.prefix === '' && (scope.get('') ?? '') !== '') declare('', '');
    return name.local;
  }
  if (name.prefix === 'xml') return `xml:${name.local}`;
  if (scope.get(name.prefix) !== name.uri) declare(name.prefix, name.uri);
  return name.prefix ? `${name.prefix}:${name.local}` : name.local;
}

function resolveAttribute(name: XmlName, scope: Map<string, string>, declare: Declare): string {
  if (name.uri === '') return name.local; // unprefixed attributes are in no namespace
  if (name.uri === XML_NS_URI) return `xml:${name.local}`;
  let prefix = name.prefix;
  if (prefix === '' || (scope.has(prefix) && scope.get(prefix) !== name.uri)) {
    // An attribute in a namespace needs a non-empty prefix bound to that namespace.
    prefix = [...scope].find(([p, u]) => p !== '' && u === name.uri)?.[0] ?? freshPrefix(scope);
  }
  if (scope.get(prefix) !== name.uri) declare(prefix, name.uri);
  return `${prefix}:${name.local}`;
}

function freshPrefix(scope: Map<string, string>): string {
  for (let i = 0; ; i++) if (!scope.has(`ns${i}`)) return `ns${i}`;
}

function writeLeaf(n: XmlNode): string {
  if (n instanceof XmlText) {
    return n.cdata ? `<![CDATA[${n.value.replaceAll(']]>', ']]]]><![CDATA[>')}]]>` : escapeText(n.value);
  }
  if (n instanceof XmlComment) return `<!--${n.value}-->`;
  if (n instanceof XmlProcessingInstruction) return n.body ? `<?${n.target} ${n.body}?>` : `<?${n.target}?>`;
  return '';
}

// XML 1.0 forbids C0 controls other than tab, LF, CR; they cannot even be written as references.
// eslint-disable-next-line no-control-regex
const INVALID_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;

export function escapeText(s: string): string {
  return s
    .replace(INVALID_CHARS, '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('\r', '&#xD;');
}

export function escapeAttr(s: string): string {
  return s
    .replace(INVALID_CHARS, '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll('\n', '&#xA;')
    .replaceAll('\r', '&#xD;')
    .replaceAll('\t', '&#x9;');
}
