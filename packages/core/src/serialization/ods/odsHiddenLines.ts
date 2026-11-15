import type { XmlElement } from '../../xml/index.ts';
import { Table } from './odsXmlNames.ts';

/**
 * Hidden and filtered rows and columns in `table:visibility`. A row a filter hid is `filter`, so
 * LibreOffice shows it again when the filter is cleared; a line hidden by hand is `collapse`.
 */

// ── Reading ─────────────────────────────────────────────────────────────────

/** The line is hidden, by hand or by a filter. */
export function isHidden(element: XmlElement): boolean {
  const visibility = element.attr(Table('visibility'));
  return visibility === 'collapse' || visibility === 'filter';
}

/** A row a filter hid (`filter`), as against one hidden by hand (`collapse`). */
export function isFiltered(element: XmlElement): boolean {
  return element.attr(Table('visibility')) === 'filter';
}

/** How a row element is hidden, in the model's terms. */
export function readRowVisibility(element: XmlElement): { filtered: boolean; manuallyHidden: boolean } {
  const filtered = isFiltered(element);
  return { filtered, manuallyHidden: isHidden(element) && !filtered };
}

// ── Writing ─────────────────────────────────────────────────────────────────

/**
 * Write a line's visibility. A row both filtered and hidden by hand is written as filtered, as
 * LibreOffice does.
 */
export function writeVisibility(element: XmlElement, line: { filtered?: boolean; hidden: boolean }): void {
  if (line.filtered) element.setAttr(Table('visibility'), 'filter');
  else if (line.hidden) element.setAttr(Table('visibility'), 'collapse');
}
