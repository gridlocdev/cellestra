import { CellStyles, NumberFormats, type CellStyle, type NumberFormat } from '../../styling/cellStyle.ts';
import type { StylePool } from '../../styling/stylePool.ts';
import type { XmlElement } from '../../xml/index.ts';
import {
  OdsCellStyles,
  parseInvariantFloat,
  unquoteFamily,
  type MapStyleTarget,
} from './mapping/odsCellStyles.ts';
import { OdsDataStyles } from './mapping/odsDataStyles.ts';
import { CM_PER_INCH, cmToPx } from './odsContentShared.ts';
import { NumberNs, Office, Style, Svg } from './odsXmlNames.ts';

/** How deep a `style:parent-style-name` chain may go before it is treated as a cycle. */
const MAX_STYLE_DEPTH = 32;

/**
 * The read side's view of every style a cell can name. `content.xml` holds the automatic styles;
 * `styles.xml` holds the common (named) styles and, in files LibreOffice wrote, many of the data
 * styles (`N129`) the automatic styles point at. Names resolve content first, then styles.xml.
 */
export class OdsStyleIndex {
  private readonly pool: StylePool;
  private readonly automaticCellStyles = new Map<string, XmlElement>();
  private readonly commonCellStyles = new Map<string, XmlElement>();
  private readonly dataStyles = new Map<string, XmlElement>();
  private readonly fontFaces = new Map<string, string>();
  private readonly columnWidths = new Map<string, number>();
  private readonly rowHeights = new Map<string, number>();
  private readonly styleIds = new Map<string, number | null>();
  private readonly formats = new Map<string, NumberFormat>();
  private plainId: number | null = null;
  /**
   * Content automatic styles the model does not regenerate (table, graphic, paragraph, text): the
   * styles preserved content (a sheet's `table:style-name`, a comment's shapes) refers to.
   */
  readonly preservedAutomaticStyles: XmlElement[] = [];

  constructor(contentRoot: XmlElement, stylesRoot: XmlElement | null, pool: StylePool) {
    this.pool = pool;
    for (const root of [stylesRoot, contentRoot]) {
      if (!root) continue;
      for (const face of root.element(Office('font-face-decls'))?.elements(Style('font-face')) ?? []) {
        const name = face.attr(Style('name'));
        const family = face.attr(Svg('font-family'));
        if (name !== null && family !== null) this.fontFaces.set(name, unquoteFamily(family));
      }
    }

    if (stylesRoot) {
      for (const container of [
        stylesRoot.element(Office('styles')),
        stylesRoot.element(Office('automatic-styles')),
      ]) {
        for (const element of container?.elements() ?? []) {
          this.indexDataStyle(element);
          if (element.is(Style('style')) && element.attr(Style('family')) === 'table-cell') {
            const name = element.attr(Style('name'));
            if (name !== null) this.commonCellStyles.set(name, element);
          }
        }
      }
    }

    for (const element of contentRoot.element(Office('automatic-styles'))?.elements() ?? []) {
      if (this.indexDataStyle(element)) continue;
      const name = element.attr(Style('name'));
      if (!element.is(Style('style')) || name === null) {
        this.preservedAutomaticStyles.push(element.clone());
        continue;
      }
      switch (element.attr(Style('family'))) {
        case 'table-cell':
          this.automaticCellStyles.set(name, element);
          break;
        case 'table-column': {
          const width = parseLength(
            element.element(Style('table-column-properties'))?.attr(Style('column-width')) ?? null,
          );
          if (width !== null) this.columnWidths.set(name, cmToPx(width));
          break;
        }
        case 'table-row': {
          const properties = element.element(Style('table-row-properties'));
          // An optimal-height row's height is whatever its content needs; the number stored with
          // it is only the last measurement (LibreOffice's 0.178in on every row), not a choice.
          if (properties?.attr(Style('use-optimal-row-height')) === 'true') break;
          const height = parseLength(properties?.attr(Style('row-height')) ?? null);
          if (height !== null) this.rowHeights.set(name, cmToPx(height));
          break;
        }
        default:
          this.preservedAutomaticStyles.push(element.clone());
      }
    }
  }

  /** The style pool id for a cell style name, or null when no such cell style exists. */
  cellStyleId(name: string): number | null {
    const cached = this.styleIds.get(name);
    if (cached !== undefined) return cached;
    const style = this.cellStyle(name, 0);
    // An automatic style is written back from its element; a common style keeps its name in
    // styles.xml and is written as an automatic style made from the model, as before.
    const element = this.automaticCellStyles.get(name);
    const id =
      style === null ? null : this.pool.intern(element ? OdsCellStyles.withSource(style, element) : style);
    this.styleIds.set(name, id);
    return id;
  }

  /**
   * A style that looks like the default but is a style of its own. A cell under a column format
   * keeps it, where the default style would show the column's format instead (as a cell that
   * names `Default` does in LibreOffice). It is written back as a plain automatic style.
   */
  plainCellStyleId(): number {
    this.plainId ??= this.pool.intern(
      OdsCellStyles.withSource(CellStyles.default, OdsCellStyles.build('Default', CellStyles.default, null)),
    );
    return this.plainId;
  }

  columnWidth(name: string): number | null {
    return this.columnWidths.get(name) ?? null;
  }

  rowHeight(name: string): number | null {
    return this.rowHeights.get(name) ?? null;
  }

  /** The number format a data style name stands for (General when it is unknown). */
  numberFormat(name: string): NumberFormat {
    let format = this.formats.get(name);
    if (!format) {
      const element = this.dataStyles.get(name);
      format = element
        ? NumberFormats.of(OdsDataStyles.toCode(element, (n) => this.dataStyles.get(n) ?? null))
        : NumberFormats.general;
      this.formats.set(name, format);
    }
    return format;
  }

  private indexDataStyle(element: XmlElement): boolean {
    if (element.uri !== NumberNs.uri) return false;
    const name = element.attr(Style('name'));
    if (name !== null) this.dataStyles.set(name, element);
    return true;
  }

  /**
   * What a `style:map`'s `style:apply-style-name` names. LibreOffice applies only common (named)
   * styles, so a common style wins over an automatic one of the same name. The style is read
   * without its own maps, so maps that name each other cannot recurse, and is interned so a save
   * that writes it as an automatic style finds it in the pool.
   */
  private mapStyle(name: string): MapStyleTarget | null {
    const common = this.commonCellStyles.has(name);
    const style = this.cellStyle(name, 0, common, false);
    if (style === null) return null;
    this.pool.intern(style);
    return { style, common };
  }

  /** Resolve a cell style through its parent chain. Automatic styles shadow common ones by name. */
  private cellStyle(name: string, depth: number, commonOnly = false, withMaps = true): CellStyle | null {
    if (depth > MAX_STYLE_DEPTH) return CellStyles.default;
    const element =
      (commonOnly ? undefined : this.automaticCellStyles.get(name)) ?? this.commonCellStyles.get(name);
    if (!element) return null;
    const parentName = element.attr(Style('parent-style-name'));
    // A parent is always a common style (ODF 16.2): an automatic style cannot be inherited from.
    const base =
      parentName === null
        ? CellStyles.default
        : (this.cellStyle(parentName, depth + 1, true, false) ?? CellStyles.default);
    return OdsCellStyles.parse(
      element,
      (n) => this.numberFormat(n),
      base,
      (n) => this.fontFaces.get(n) ?? null,
      withMaps ? (n) => this.mapStyle(n) : null,
    );
  }
}

/** A length in cm or inches → cm; null for anything else. */
export function parseLength(value: string | null): number | null {
  if (value === null) return null;
  const lower = value.trim().toLowerCase();
  if (lower.endsWith('cm')) return parseInvariantFloat(lower.slice(0, -2));
  if (lower.endsWith('mm')) {
    const mm = parseInvariantFloat(lower.slice(0, -2));
    return mm === null ? null : mm / 10;
  }
  if (lower.endsWith('in')) {
    const inches = parseInvariantFloat(lower.slice(0, -2));
    return inches === null ? null : inches * CM_PER_INCH;
  }
  if (lower.endsWith('pt')) {
    const pt = parseInvariantFloat(lower.slice(0, -2));
    return pt === null ? null : (pt / 72) * CM_PER_INCH;
  }
  return null;
}
