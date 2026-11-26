/**
 * Regions for F6 and Shift+F6: the big parts of the window the keyboard jumps between (the ribbon
 * tabs or the compact toolbar, the grid, the sheet tabs). Each one carries `REGION`; the order is
 * the document's. Inside a region, focus goes to its current item (the `data-focused` one that
 * roving focus keeps in app state), so arriving puts you where you left it.
 */

/** The region's attribute: put it on the element that holds the region. */
export function region(name: string): { readonly 'data-region': string } {
  return { 'data-region': name };
}

/** Just what `nextRegion` needs of a DOM node, so it can be tested without a document. */
export interface RegionNode {
  contains(other: RegionNode | null): boolean;
  compareDocumentPosition(other: RegionNode): number;
}

const FOLLOWING = 4; // Node.DOCUMENT_POSITION_FOLLOWING

/**
 * The region F6 (`step` 1) or Shift+F6 (`step` -1) goes to from `active`, among `regions` in
 * document order. From inside a region it steps and wraps. From outside every region (the formula
 * bar, nothing focused) it goes to the first region after the focus, or the last one before it.
 */
export function nextRegion<T extends RegionNode>(
  regions: readonly T[],
  active: RegionNode | null,
  step: 1 | -1,
): T | undefined {
  const count = regions.length;
  if (count === 0) return undefined;
  const inside = active ? regions.findIndex((r) => r.contains(active)) : -1;
  if (inside >= 0) return regions[(inside + step + count) % count];
  if (active) {
    const after = (r: T) => (active.compareDocumentPosition(r) & FOLLOWING) !== 0;
    const found = step === 1 ? regions.find(after) : regions.findLast((r) => !after(r));
    if (found) return found;
  }
  return regions[step === 1 ? 0 : count - 1];
}

/** The element in a region that takes focus: its current item, or else its first Tab stop. */
function focusTarget(el: HTMLElement): HTMLElement | null {
  return (
    el.querySelector<HTMLElement>('[data-focused]') ??
    (el.tabIndex >= 0 ? el : null) ??
    el.querySelector<HTMLElement>('[tabindex="0"], button:not([disabled]):not([tabindex="-1"])')
  );
}

/** Moves focus to the next (1) or previous (-1) region on screen. False if there is none. */
export function focusRegion(doc: Document, step: 1 | -1): boolean {
  const shown = [...doc.querySelectorAll<HTMLElement>('[data-region]')].filter(
    (el) => el.getClientRects().length > 0 && focusTarget(el) !== null,
  );
  const to = nextRegion(shown, doc.activeElement, step);
  const target = to && focusTarget(to);
  if (!target) return false;
  target.focus();
  return true;
}
