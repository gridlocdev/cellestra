import { CellStyles, styleKey, type CellStyle } from './cellStyle.ts';

/** Stores each distinct style once and hands out integer ids; id 0 is always the default style. */
export class StylePool {
  static readonly defaultId = 0;
  private readonly list: CellStyle[] = [];
  private readonly index = new Map<string, number>();

  constructor() {
    this.intern(CellStyles.default);
  }

  get count(): number {
    return this.list.length;
  }

  get styles(): readonly CellStyle[] {
    return this.list;
  }

  intern(style: CellStyle): number {
    const key = styleKey(style);
    const existing = this.index.get(key);
    if (existing !== undefined) return existing;
    const id = this.list.length;
    this.list.push(style);
    this.index.set(key, id);
    return id;
  }

  /** Out-of-range ids read as the default style. */
  get(id: number): CellStyle {
    return this.list[id] ?? CellStyles.default;
  }

  mutate(baseId: number, change: (s: CellStyle) => CellStyle): number {
    return this.intern(change(this.get(baseId)));
  }
}
