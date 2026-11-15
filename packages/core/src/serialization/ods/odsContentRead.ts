import type { Workbook } from '../../model/workbook.ts';
import {
  appendXmlText,
  openXmlElement,
  sameName,
  serializeXmlElement,
  XmlComment,
  XmlElement,
  XmlEventReader,
  XmlProcessingInstruction,
  type XmlAttribute,
  type XmlDocument,
  type XmlEventHandler,
  type XmlName,
} from '../../xml/index.ts';
import { applyCheckboxes, type ReadCheckboxes } from './odsCheckboxes.ts';
import { readDatabaseRanges } from './odsDatabaseRanges.ts';
import { readNamedExpressions } from './odsNamedExpressions.ts';
import type { ForeignXml, OdsPackage } from './odsPackage.ts';
import { OdsPreserved } from './odsPreserved.ts';
import type { RepeatBudget } from './odsReadLimits.ts';
import { OdsStyleIndex } from './odsStyleIndex.ts';
import { TableReader, type OpenSpan, type ReadContext } from './odsTableRead.ts';
import { assembleValidations, readContentValidationRules, type ValidationCells } from './odsValidation.ts';
import { OdsXmlNames as N, Office, Table } from './odsXmlNames.ts';

/** The row and column containers ODF allows a grid to be nested in (ODF 9.1.2, 9.1.6). */
const ROW_GROUPS: readonly XmlName[] = [
  Table('table-header-rows'),
  Table('table-row-group'),
  Table('table-rows'),
];
const COLUMN_GROUPS: readonly XmlName[] = [
  Table('table-header-columns'),
  Table('table-column-group'),
  Table('table-columns'),
];

/**
 * Read `content.xml`, given as text in one or more chunks, into `workbook`, and return the
 * document's outline for the page setup: the root with its automatic styles, and each table
 * element with its attributes but none of its content (null when there is no root element).
 *
 * The document is read as a stream: each row element is built, read into the sheet and dropped,
 * so a large file never has its whole tree in memory. Unmodeled `office:spreadsheet` children are
 * appended to `foreign` as verbatim XML (with whether they came before the first table) so the
 * writer can re-emit them in place; unmodeled table, cell and named-expression content is kept
 * through `OdsPreserved`. `stylesDocument` is styles.xml, whose common and data styles cells may
 * name. Throws `XmlParseError` when the text is not well-formed.
 */
export function readContentXml(
  chunks: Iterable<string>,
  workbook: Workbook,
  foreign: ForeignXml[],
  pkg: OdsPackage | null = null,
  stylesDocument: XmlDocument | null = null,
): XmlElement {
  const content = new ContentReader(workbook, foreign, pkg, stylesDocument?.root ?? null);
  const reader = new XmlEventReader(content);
  for (const chunk of chunks) reader.write(chunk);
  reader.close();
  return content.finish();
}

/** The row and column groups' leaves and groups, which make up a table's grid. */
const isRowGroup = (name: XmlName): boolean => ROW_GROUPS.some((g) => sameName(g, name));
const isColumnGroup = (name: XmlName): boolean => COLUMN_GROUPS.some((g) => sameName(g, name));
const TABLE_ROW = Table('table-row');
const TABLE_COLUMN = Table('table-column');

/**
 * Where the reader is in the document. `element` frames build a tree: the outline (`root`,
 * `body`, `spreadsheet`) or a `subtree` that is handed to `done` when it closes. The other
 * frames build nothing: a table, a row or column group in it, or content that is not read.
 */
type Frame =
  | {
      readonly kind: 'element';
      readonly element: XmlElement;
      readonly role: 'root' | 'body' | 'spreadsheet' | 'subtree';
      readonly done: ((e: XmlElement) => void) | null;
    }
  | { readonly kind: 'table'; readonly table: TableReader }
  | {
      readonly kind: 'rows' | 'columns';
      readonly table: TableReader;
      /** The group or print-title span this container opens; null for a plain `table:table-rows`. */
      readonly span: OpenSpan | null;
    }
  | { readonly kind: 'skip' };

const SKIP: Frame = { kind: 'skip' };

/**
 * Turns the events of `content.xml` into the workbook. The outline before `office:body`
 * (automatic styles, font faces) is built as a tree; in the first `office:spreadsheet`, each
 * child other than a table is built and read when it closes, and each table is read by a
 * `TableReader` one row element at a time.
 */
class ContentReader implements XmlEventHandler {
  private readonly workbook: Workbook;
  private readonly foreign: ForeignXml[];
  private readonly pkg: OdsPackage | null;
  private readonly stylesRoot: XmlElement | null;
  private readonly frames: Frame[] = [];
  private root: XmlElement | null = null;
  private sawBody = false;
  private spreadsheet: XmlElement | null = null;
  private styleIndex: OdsStyleIndex | null = null;
  private seenTable = false;
  private validations: XmlElement | null = null;
  private readonly validationCells: ValidationCells = new Map();
  private readonly namedExpressions: XmlElement[] = [];
  private readonly checkboxes: ReadCheckboxes = { ids: new Set(), cells: [] };
  private readonly repeats: RepeatBudget = { cells: 0, merges: 0 };

  constructor(
    workbook: Workbook,
    foreign: ForeignXml[],
    pkg: OdsPackage | null,
    stylesRoot: XmlElement | null,
  ) {
    this.workbook = workbook;
    this.foreign = foreign;
    this.pkg = pkg;
    this.stylesRoot = stylesRoot;
  }

  /** The styles cells name: content.xml's automatic styles come before its body, so are complete. */
  private get styles(): OdsStyleIndex {
    this.styleIndex ??= new OdsStyleIndex(this.root!, this.stylesRoot, this.workbook.styles);
    return this.styleIndex;
  }

  declaration(): void {}

  openElement(name: XmlName, attributes: XmlAttribute[]): void {
    const top = this.frames.at(-1);
    if (!top) {
      this.root = openXmlElement(null, name, attributes);
      this.frames.push({ kind: 'element', element: this.root, role: 'root', done: null });
      return;
    }
    switch (top.kind) {
      case 'skip':
        this.frames.push(SKIP);
        return;
      case 'rows':
        if (sameName(name, TABLE_ROW)) this.subtree(name, attributes, (e) => top.table.row(e));
        else
          this.frames.push(isRowGroup(name) ? this.lineContainer('rows', top.table, name, attributes) : SKIP);
        return;
      case 'columns':
        if (sameName(name, TABLE_COLUMN)) this.subtree(name, attributes, (e) => top.table.column(e));
        else
          this.frames.push(
            isColumnGroup(name) ? this.lineContainer('columns', top.table, name, attributes) : SKIP,
          );
        return;
      case 'table':
        this.openInTable(top.table, name, attributes);
        return;
      case 'element':
        this.openInOutline(top, name, attributes);
    }
  }

  private openInOutline(top: Frame & { kind: 'element' }, name: XmlName, attributes: XmlAttribute[]): void {
    switch (top.role) {
      case 'root': {
        const element = openXmlElement(top.element, name, attributes);
        const body = !this.sawBody && sameName(name, Office('body'));
        if (body) this.sawBody = true;
        this.frames.push({ kind: 'element', element, role: body ? 'body' : 'subtree', done: null });
        return;
      }
      case 'body':
        // Only the first office:spreadsheet of the body is read.
        if (this.spreadsheet || !sameName(name, Office('spreadsheet'))) {
          this.frames.push(SKIP);
          return;
        }
        this.spreadsheet = openXmlElement(top.element, name, attributes);
        this.frames.push({ kind: 'element', element: this.spreadsheet, role: 'spreadsheet', done: null });
        return;
      case 'spreadsheet':
        if (sameName(name, Table('table'))) {
          this.seenTable = true;
          // The outline keeps the table's attributes, which the page setup reads.
          const outline = openXmlElement(top.element, name, attributes);
          this.frames.push({ kind: 'table', table: this.startTable(outline) });
        } else {
          this.subtree(name, attributes, (e) => this.spreadsheetChild(e));
        }
        return;
      case 'subtree': {
        const element = openXmlElement(top.element, name, attributes);
        this.frames.push({ kind: 'element', element, role: 'subtree', done: null });
      }
    }
  }

  private openInTable(table: TableReader, name: XmlName, attributes: XmlAttribute[]): void {
    if (sameName(name, TABLE_ROW)) {
      table.gridStarted();
      this.subtree(name, attributes, (e) => table.row(e));
    } else if (sameName(name, TABLE_COLUMN)) {
      table.gridStarted();
      this.subtree(name, attributes, (e) => table.column(e));
    } else if (isRowGroup(name)) {
      table.gridStarted();
      this.frames.push(this.lineContainer('rows', table, name, attributes));
    } else if (isColumnGroup(name)) {
      table.gridStarted();
      this.frames.push(this.lineContainer('columns', table, name, attributes));
    } else {
      const afterGrid = table.hasGrid;
      this.subtree(name, attributes, (e) => table.child(e, afterGrid));
    }
  }

  /**
   * A frame for a row or column container. A group (`table:table-row-group`) and the print titles
   * (`table:table-header-rows`) are modeled spans, closed when the container closes; a plain
   * `table:table-rows` says nothing about its lines.
   */
  private lineContainer(
    kind: 'rows' | 'columns',
    table: TableReader,
    name: XmlName,
    attributes: XmlAttribute[],
  ): Frame {
    const header = name.local === `table-header-${kind}`;
    const group = name.local === (kind === 'rows' ? 'table-row-group' : 'table-column-group');
    if (!header && !group) return { kind, table, span: null };
    const display = attributes.find((a) => a.name.uri === Table.uri && a.name.local === 'display');
    return {
      kind,
      table,
      span: { header, first: table.position(kind), collapsed: display?.value === 'false' },
    };
  }

  /** Build a detached element and everything in it, and hand it to `done` when it closes. */
  private subtree(name: XmlName, attributes: XmlAttribute[], done: (e: XmlElement) => void): void {
    const element = openXmlElement(null, name, attributes);
    this.frames.push({ kind: 'element', element, role: 'subtree', done });
  }

  closeElement(): void {
    const frame = this.frames.pop();
    if (frame?.kind === 'element') {
      frame.element.trimToSize();
      frame.done?.(frame.element);
    } else if (frame?.kind === 'table') {
      frame.table.finish();
    } else if ((frame?.kind === 'rows' || frame?.kind === 'columns') && frame.span) {
      frame.table.closeSpan(frame.kind, frame.span);
    }
  }

  text(value: string, cdata: boolean): void {
    const top = this.frames.at(-1);
    if (top?.kind === 'element') appendXmlText(top.element, value, cdata, true);
  }

  comment(value: string): void {
    const top = this.frames.at(-1);
    if (top?.kind === 'element') top.element.nodes.push(new XmlComment(value));
  }

  processingInstruction(target: string, body: string): void {
    const top = this.frames.at(-1);
    if (top?.kind === 'element') top.element.nodes.push(new XmlProcessingInstruction(target, body));
  }

  private startTable(outline: XmlElement): TableReader {
    const workbook = this.workbook;
    const name = outline.attr(Table('name')) ?? workbook.addSheet().name;
    const sheet = workbook.findSheet(name) ?? workbook.addSheetForLoad(name);
    const context: ReadContext = {
      styles: this.styles,
      validationCells: this.validationCells,
      chartSheet: false,
      checkboxIds: this.checkboxes.ids,
      merges: [],
      repeats: this.repeats,
    };
    return new TableReader(workbook, sheet, outline, context, this.pkg, this.checkboxes);
  }

  private spreadsheetChild(child: XmlElement): void {
    if (child.is(Table('named-expressions'))) {
      this.namedExpressions.push(...readNamedExpressions(child, this.workbook, null));
    } else if (child.is(Table(N.databaseRanges))) {
      readDatabaseRanges(child, this.workbook);
    } else if (child.is(Table(N.contentValidations))) {
      // Modeled: the rules are read at the end, once every cell has named the ones it uses.
      this.validations ??= child;
    } else {
      this.foreign.push({ beforeTables: !this.seenTable, xml: serializeXmlElement(child) });
    }
  }

  /** The document is complete: settle what needed every table, and return the outline. */
  finish(): XmlElement {
    const root = this.root!;
    if (!this.spreadsheet) return root;
    // Rule definitions are read last; each rule's target is rebuilt from the cells that named it.
    const container = new XmlElement(Office('spreadsheet'), this.validations);
    assembleValidations(this.workbook, readContentValidationRules(container), this.validationCells);
    applyCheckboxes(this.workbook, this.checkboxes);
    OdsPreserved.setWorkbook(this.workbook, {
      spreadsheetAttributes: this.spreadsheet.attributes.map((a) => ({ name: a.name, value: a.value })),
      namedExpressions: this.namedExpressions,
      automaticStyles: this.styles.preservedAutomaticStyles,
      scripts: root.element(Office('scripts'))?.clone() ?? null,
      fontFaces: root.element(Office('font-face-decls'))?.clone() ?? null,
    });
    return root;
  }
}
