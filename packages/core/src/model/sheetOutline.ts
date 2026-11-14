/** A run of whole rows or columns, `first` … `last` inclusive (0-based). */
export interface LineSpan {
  readonly first: number;
  readonly last: number;
}

/** A row or column group (LibreOffice's Data ▸ Group and Outline); collapsed when folded away. */
export interface LineGroup extends LineSpan {
  readonly collapsed: boolean;
}

/**
 * A sheet's row and column groups and its print titles. Cellestra has no command that changes
 * them; they are read from a file, move with their lines on insert and delete, and are saved back
 * (PRD §8.1). Groups nest: an inner group lies within its outer one.
 */
export interface SheetOutline {
  readonly rowGroups: readonly LineGroup[];
  readonly columnGroups: readonly LineGroup[];
  /** Rows printed at the top of every page (ODF `table:table-header-rows`); null when none. */
  readonly headerRows: LineSpan | null;
  /** Columns printed at the left of every page (ODF `table:table-header-columns`); null when none. */
  readonly headerColumns: LineSpan | null;
}

export const SheetOutline = {
  empty: { rowGroups: [], columnGroups: [], headerRows: null, headerColumns: null } as SheetOutline,

  isEmpty(outline: SheetOutline): boolean {
    return (
      outline.rowGroups.length === 0 &&
      outline.columnGroups.length === 0 &&
      outline.headerRows === null &&
      outline.headerColumns === null
    );
  },
};
