import { NumberFormats, type CellStyleMap, type NumberFormat } from '../../styling/cellStyle.ts';
import type { StylePool } from '../../styling/stylePool.ts';
import { XmlElement } from '../../xml/index.ts';
import { OdsCellStyles, type WrittenStyleMap } from './mapping/odsCellStyles.ts';
import { OdsDataStyles } from './mapping/odsDataStyles.ts';
import { formatUpTo3, pxToCm } from './odsContentShared.ts';
import { Style } from './odsXmlNames.ts';

/** Write side: interns cell/data/column/row automatic styles into `office:automatic-styles`. */
export class StyleRegistry {
  readonly styles: StylePool;
  private readonly autoStyles: XmlElement;
  private readonly cellStyleNames = new Map<number, string>();
  private readonly dataStyleNames = new Map<string, string>();
  private readonly columnStyleNames = new Map<number, string>();
  private readonly rowStyleNames = new Map<number, string>();
  private next = 0;

  /** Names already used by preserved automatic styles and by styles.xml, which generated names must not take. */
  private readonly reserved: ReadonlySet<string>;
  /** The names styles.xml declares, which is saved as it was: a `style:map` can keep naming one. */
  private readonly commonNames: ReadonlySet<string>;

  constructor(
    styles: StylePool,
    autoStyles: XmlElement,
    reserved: ReadonlySet<string> = new Set(),
    commonNames: ReadonlySet<string> = new Set(),
  ) {
    this.styles = styles;
    this.autoStyles = autoStyles;
    this.reserved = reserved;
    this.commonNames = commonNames;
  }

  private nextName(prefix: string): string {
    for (;;) {
      const name = `${prefix}${this.next++}`;
      // A sectioned data style also names its sections `N3P0`, `N3P1`, … (as LibreOffice does).
      const sections = prefix === 'N' ? [0, 1, 2].map((i) => `${name}P${i}`) : [];
      if (!this.reserved.has(name) && !sections.some((n) => this.reserved.has(n))) return name;
    }
  }

  cellStyle(styleId: number): string {
    const existing = this.cellStyleNames.get(styleId);
    if (existing !== undefined) return existing;

    const style = this.styles.get(styleId);
    const dataStyleName = NumberFormats.isGeneral(style.numberFormat)
      ? null
      : this.dataStyle(style.numberFormat);

    // A style a map applies is written (and named) first, so the map can point at it.
    const maps = (style.conditionalMaps ?? []).flatMap((map): WrittenStyleMap[] => {
      const applyStyleName = this.mapTarget(map);
      return applyStyleName === null
        ? []
        : [{ condition: map.condition, applyStyleName, baseCellAddress: map.baseCellAddress }];
    });

    const name = this.nextName('ce');
    this.autoStyles.add(OdsCellStyles.build(name, style, dataStyleName, maps));
    this.cellStyleNames.set(styleId, name);
    return name;
  }

  /**
   * The name a `style:map` applies on save: the common style it named while styles.xml still
   * declares it (or when nothing else is known), otherwise an automatic style written from what it
   * resolved to on read, under this save's name for it.
   */
  private mapTarget(map: CellStyleMap): string | null {
    if (
      map.applyStyleName !== null &&
      (map.applyStyle === null || this.commonNames.has(map.applyStyleName))
    ) {
      return map.applyStyleName;
    }
    return map.applyStyle === null ? null : this.cellStyle(this.styles.intern(map.applyStyle));
  }

  private dataStyle(format: NumberFormat): string | null {
    const existing = this.dataStyleNames.get(format.code);
    if (existing !== undefined) return existing;

    const name = this.nextName('N');
    // A sectioned code is several elements: the per-section styles first, then the one named `name`.
    const elements = OdsDataStyles.build(name, format);
    // A code no data style can express: the cell writes no data-style-name rather than one that
    // names nothing.
    if (elements.length === 0) return null;

    for (const element of elements) this.autoStyles.add(element);
    this.dataStyleNames.set(format.code, name);
    return name;
  }

  columnStyle(width: number): string {
    const existing = this.columnStyleNames.get(width);
    if (existing !== undefined) return existing;

    const name = this.nextName('co');
    this.autoStyles.add(
      new XmlElement(Style('style'))
        .setAttr(Style('name'), name)
        .setAttr(Style('family'), 'table-column')
        .add(
          new XmlElement(Style('table-column-properties')).setAttr(
            Style('column-width'),
            formatCm(pxToCm(width)),
          ),
        ),
    );
    this.columnStyleNames.set(width, name);
    return name;
  }

  rowStyle(height: number): string {
    const existing = this.rowStyleNames.get(height);
    if (existing !== undefined) return existing;

    const name = this.nextName('ro');
    this.autoStyles.add(
      new XmlElement(Style('style'))
        .setAttr(Style('name'), name)
        .setAttr(Style('family'), 'table-row')
        .add(
          new XmlElement(Style('table-row-properties'))
            .setAttr(Style('row-height'), formatCm(pxToCm(height)))
            // A height someone set is kept, not re-measured to fit the content when LibreOffice loads it.
            .setAttr(Style('use-optimal-row-height'), 'false'),
        ),
    );
    this.rowStyleNames.set(height, name);
    return name;
  }
}

function formatCm(cm: number): string {
  return formatUpTo3(cm) + 'cm';
}
