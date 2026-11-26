/**
 * Roving focus for tab strips, menus and lists. The current item lives in app state (a signal or
 * component state), not in the DOM: the group renders it with `rovingItem(true)` and every other item
 * with `rovingItem(false)`. Only the current item is a Tab stop, it carries `data-focused` (which the
 * stylesheet draws in --cs-focus while the group has focus), and DOM focus follows the state when it
 * changes. A key moves the state through `rovingStep`; the focus you see follows by itself.
 */

export type RovingAxis = 'horizontal' | 'vertical';

type Key = Pick<KeyboardEvent, 'key' | 'shiftKey' | 'ctrlKey' | 'metaKey' | 'altKey'>;

/**
 * Where a key moves the current item in a group of `count`: the arrows on the group's axis step and
 * wrap, Home and End go to the ends. Items `skip` names (disabled ones) are stepped over. `undefined`
 * for any other key, or a chord with ⌘, Ctrl or ⌥.
 */
export function rovingStep(
  e: Key,
  index: number,
  count: number,
  axis: RovingAxis,
  skip: (i: number) => boolean = () => false,
): number | undefined {
  if (count <= 0 || e.ctrlKey || e.metaKey || e.altKey) return undefined;
  const [back, forward] = axis === 'horizontal' ? ['ArrowLeft', 'ArrowRight'] : ['ArrowUp', 'ArrowDown'];
  const walk = (from: number, step: 1 | -1): number | undefined => {
    for (let n = 0, i = from; n < count; n++, i = (i + step + count) % count) {
      if (!skip(i)) return i;
    }
    return undefined;
  };
  switch (e.key) {
    case forward:
      return walk((index + 1 + count) % count, 1);
    case back:
      return walk((index - 1 + count) % count, -1);
    case 'Home':
      return walk(0, 1);
    case 'End':
      return walk(count - 1, -1);
    default:
      return undefined;
  }
}

/** The group's attribute: put it on the element that holds the items. */
export const ROVING_GROUP = { 'data-roving': '' } as const;

/**
 * A ref for the current item: when the state moves to it while focus is already in its group (a key
 * moved it), it takes DOM focus. A change from elsewhere (a click on the grid, a command, an undo)
 * moves the mark but leaves focus where it is.
 */
export function followFocus(el: HTMLElement | null): void {
  if (!el) return;
  const group = el.closest('[data-roving]');
  const active = el.ownerDocument.activeElement;
  if (group && active && active !== el && group.contains(active)) el.focus();
}

/** A ref for the current item of a menu or list that takes the keyboard as it opens. */
export function takeFocus(el: HTMLElement | null): void {
  if (el && el.ownerDocument.activeElement !== el) el.focus();
}

/** Props for one item: a Tab stop and `data-focused` only for the current item, and its focus ref. */
export function rovingItem(
  current: boolean,
  ref: (el: HTMLElement | null) => void = followFocus,
): { tabIndex: number; 'data-focused'?: ''; ref?: (el: HTMLElement | null) => void } {
  return current ? { tabIndex: 0, 'data-focused': '', ref } : { tabIndex: -1 };
}
