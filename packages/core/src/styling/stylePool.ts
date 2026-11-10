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

  /**
   * Put `style` at `id`, keeping the id, so every cell and line using it changes with it (a
   * structural edit moving a conditional map's base cell). Another id may then hold an equal style;
   * `intern` keeps answering with whichever it already knew.
   */
  replace(id: number, style: CellStyle): void {
    const old = this.list[id];
    if (old === undefined || id === StylePool.defaultId) return;
    const oldKey = styleKey(old);
    if (this.index.get(oldKey) === id) this.index.delete(oldKey);
    this.list[id] = style;
    const key = styleKey(style);
    if (!this.index.has(key)) this.index.set(key, id);
  }

  mutate(baseId: number, change: (s: CellStyle) => CellStyle): number {
    return this.intern(change(this.get(baseId)));
  }
}
