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

  constructor(styles: StylePool, autoStyles: XmlElement) {
    this.styles = styles;
    this.autoStyles = autoStyles;
  }

  cellStyle(styleId: number): string {
    const existing = this.cellStyleNames.get(styleId);
    if (existing !== undefined) return existing;

    const style = this.styles.get(styleId);
    const dataStyleName = NumberFormats.isGeneral(style.numberFormat)
      ? null
      : this.dataStyle(style.numberFormat);

    const name = `ce${this.next++}`;
    this.autoStyles.add(OdsCellStyles.build(name, style, dataStyleName));
    this.cellStyleNames.set(styleId, name);
    return name;
  }

  private dataStyle(format: NumberFormat): string {
    const existing = this.dataStyleNames.get(format.code);
    if (existing !== undefined) return existing;

    const name = `N${this.next++}`;
    const element = OdsDataStyles.build(name, format);
    if (!element) return name; // General; should not happen since caller checks

    this.autoStyles.add(element);
    this.dataStyleNames.set(format.code, name);
    return name;
  }

  columnStyle(width: number): string {
    const existing = this.columnStyleNames.get(width);
    if (existing !== undefined) return existing;

    const name = `co${this.next++}`;
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

    const name = `ro${this.next++}`;
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
