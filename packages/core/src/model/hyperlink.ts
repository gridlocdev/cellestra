/** A hyperlink on a cell (web or file URL), ↔ ODF `text:a`. The visible text is the cell value. */
export interface Hyperlink {
  readonly target: string;
  readonly tooltip?: string | null;
}
