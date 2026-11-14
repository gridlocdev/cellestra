import type { LineGroup, LineSpan } from '../../model/sheetOutline.ts';
import { XmlElement } from '../../xml/index.ts';
import { Table } from './odsXmlNames.ts';

interface Span {
  readonly first: number;
  readonly last: number;
  readonly header: boolean;
  readonly collapsed: boolean;
}

interface Open {
  readonly last: number;
  readonly header: boolean;
  readonly element: XmlElement;
}

/**
 * Writes a sheet's row (or column) groups and print titles as the containers ODF nests its grid
 * lines in: `table:table-row-group` (with `table:display="false"` when collapsed) and
 * `table:table-header-rows`. The grid writer asks where each line goes; a repeated run must not
 * cross a container's edge, so it also breaks its runs at `boundaries`.
 */
export class LineContainers {
  private readonly table: XmlElement;
  private readonly axis: 'row' | 'column';
  private readonly spans: Span[];
  private readonly stack: Open[] = [];
  private next = 0;

  constructor(
    table: XmlElement,
    axis: 'row' | 'column',
    groups: readonly LineGroup[],
    header: LineSpan | null,
  ) {
    this.table = table;
    this.axis = axis;
    const spans: Span[] = groups.map((g) => ({ ...g, header: false }));
    if (header) spans.push({ ...header, header: true, collapsed: false });
    // Outer before inner: earlier first, then longer, then a group around the print titles.
    this.spans = spans.sort(
      (a, b) => a.first - b.first || b.last - a.last || Number(a.header) - Number(b.header),
    );
  }

  /** The last line any container covers (-1 for none): the grid is written at least that far. */
  get lastLine(): number {
    return this.spans.reduce((last, s) => Math.max(last, s.last), -1);
  }

  /** The lines where a container starts, or the line after one ends. */
  get boundaries(): number[] {
    return this.spans.flatMap((s) => [s.first, s.last + 1]);
  }

  /** The element a line starting at `position` is added to, opening and closing containers on the way. */
  at(position: number): XmlElement {
    while (this.stack.length > 0 && this.stack.at(-1)!.last < position) this.stack.pop();
    while (this.next < this.spans.length && this.spans[this.next]!.first <= position) {
      const span = this.spans[this.next++]!;
      const parent = this.stack.at(-1);
      // Print titles hold lines only, and a container must lie within the one around it.
      if (span.last < position || parent?.header) continue;
      const element = this.element(span);
      (parent?.element ?? this.table).add(element);
      this.stack.push({ last: Math.min(span.last, parent?.last ?? span.last), header: span.header, element });
    }
    return this.stack.at(-1)?.element ?? this.table;
  }

  private element(span: Span): XmlElement {
    if (span.header) return new XmlElement(Table(`table-header-${this.axis}s`));
    const element = new XmlElement(Table(`table-${this.axis}-group`));
    if (span.collapsed) element.setAttr(Table('display'), 'false');
    return element;
  }
}
