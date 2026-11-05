import { NumberFormats, type NumberFormat } from '../../styling/cellStyle.ts';
import type { StylePool } from '../../styling/stylePool.ts';
import { XmlElement } from '../../xml/index.ts';
import { OdsCellStyles } from './mapping/odsCellStyles.ts';
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

  /** Names already used by preserved automatic styles, which generated names must not take. */
  private readonly reserved: ReadonlySet<string>;

  constructor(styles: StylePool, autoStyles: XmlElement, reserved: ReadonlySet<string> = new Set()) {
    this.styles = styles;
    this.autoStyles = autoStyles;
    this.reserved = reserved;
  }

  private nextName(prefix: string): string {
    for (;;) {
      const name = `${prefix}${this.next++}`;
      if (!this.reserved.has(name)) return name;
    }
  }

  cellStyle(styleId: number): string {
    const existing = this.cellStyleNames.get(styleId);
    if (existing !== undefined) return existing;

    const style = this.styles.get(styleId);
    const dataStyleName = NumberFormats.isGeneral(style.numberFormat)
      ? null
      : this.dataStyle(style.numberFormat);

    const name = this.nextName('ce');
    this.autoStyles.add(OdsCellStyles.build(name, style, dataStyleName));
    this.cellStyleNames.set(styleId, name);
    return name;
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
          new XmlElement(Style('table-row-properties')).setAttr(
            Style('row-height'),
            formatCm(pxToCm(height)),
          ),
        ),
    );
    this.rowStyleNames.set(height, name);
    return name;
  }
}

function formatCm(cm: number): string {
  return formatUpTo3(cm) + 'cm';
}
