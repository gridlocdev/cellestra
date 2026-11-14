/**
 * What the formula bar's field holds, and whether it is the user's typing. Focus alone is not an
 * edit: until the user types, the field follows the active cell. Typing is tied to the cell it
 * started in (`target`), so a sheet switch or a new document drops it instead of writing it into
 * whatever cell is active by then.
 */
export class FormulaBarDraft {
  text: string;
  #dirty = false;
  #target: string | null = null;

  constructor(content: string) {
    this.text = content;
  }

  get dirty(): boolean {
    return this.#dirty;
  }

  /** The user typed (or pasted, or added a line break). */
  input(text: string, target: string): void {
    if (!this.#dirty) this.#target = target;
    this.#dirty = true;
    this.text = text;
  }

  /** The active cell's content or identity changed. */
  sync(content: string, target: string): void {
    if (this.#dirty && target === this.#target) return;
    this.reset(content);
  }

  /** Escape, or the edit was committed. */
  reset(content: string): void {
    this.#dirty = false;
    this.#target = null;
    this.text = content;
  }

  /** The text to commit when the field loses focus or Enter is pressed, or null for nothing. */
  take(content: string): string | null {
    const text = this.#dirty && this.text !== content ? this.text : null;
    this.reset(text ?? content);
    return text;
  }
}
