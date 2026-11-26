import { describe, expect, it } from 'vitest';
import { followFocus, rovingItem, rovingStep, takeFocus } from '../src/components/roving.ts';

const key = (
  k: string,
  mods: Partial<Record<'shiftKey' | 'ctrlKey' | 'metaKey' | 'altKey', boolean>> = {},
) => ({
  key: k,
  shiftKey: false,
  ctrlKey: false,
  metaKey: false,
  altKey: false,
  ...mods,
});

describe('rovingStep', () => {
  it('steps along its axis and wraps at both ends', () => {
    expect(rovingStep(key('ArrowRight'), 0, 3, 'horizontal')).toBe(1);
    expect(rovingStep(key('ArrowRight'), 2, 3, 'horizontal')).toBe(0);
    expect(rovingStep(key('ArrowLeft'), 0, 3, 'horizontal')).toBe(2);
    expect(rovingStep(key('ArrowDown'), 1, 3, 'vertical')).toBe(2);
    expect(rovingStep(key('ArrowUp'), 0, 3, 'vertical')).toBe(2);
  });

  it('ignores the other axis, other keys and chords', () => {
    expect(rovingStep(key('ArrowDown'), 0, 3, 'horizontal')).toBeUndefined();
    expect(rovingStep(key('ArrowRight'), 0, 3, 'vertical')).toBeUndefined();
    expect(rovingStep(key('a'), 0, 3, 'vertical')).toBeUndefined();
    expect(rovingStep(key('ArrowDown', { metaKey: true }), 0, 3, 'vertical')).toBeUndefined();
    expect(rovingStep(key('ArrowDown', { ctrlKey: true }), 0, 3, 'vertical')).toBeUndefined();
    expect(rovingStep(key('ArrowDown', { altKey: true }), 0, 3, 'vertical')).toBeUndefined();
  });

  it('goes to the ends with Home and End', () => {
    expect(rovingStep(key('Home'), 2, 4, 'vertical')).toBe(0);
    expect(rovingStep(key('End'), 0, 4, 'vertical')).toBe(3);
  });

  it('steps over skipped items, and has nowhere to go when all are skipped or the group is empty', () => {
    const disabled = (i: number) => i === 1 || i === 3;
    expect(rovingStep(key('ArrowDown'), 0, 4, 'vertical', disabled)).toBe(2);
    expect(rovingStep(key('ArrowUp'), 0, 4, 'vertical', disabled)).toBe(2);
    expect(rovingStep(key('End'), 0, 4, 'vertical', disabled)).toBe(2);
    expect(rovingStep(key('Home'), 2, 4, 'vertical', (i) => i === 0)).toBe(1);
    expect(rovingStep(key('ArrowDown'), 0, 2, 'vertical', () => true)).toBeUndefined();
    expect(rovingStep(key('ArrowDown'), 0, 0, 'vertical')).toBeUndefined();
  });
});

/** Just enough of an element for the focus refs: a group, its items and the document's active element. */
function fakeGroup() {
  const doc = { activeElement: null as unknown };
  const group = { contains: (el: unknown) => items.includes(el as never) };
  const make = () => {
    const el = {
      ownerDocument: doc,
      closest: (selector: string) => (selector === '[data-roving]' ? group : null),
      focus: () => {
        doc.activeElement = el;
      },
    };
    return el;
  };
  const items = [make(), make()];
  const outside = make();
  outside.closest = () => null;
  return { doc, items, outside };
}

describe('roving focus refs', () => {
  it('moves focus to the current item only when focus is already in its group', () => {
    const { doc, items, outside } = fakeGroup();
    const [a, b] = items as [HTMLElement & { focus(): void }, HTMLElement];
    outside.focus();
    followFocus(b);
    expect(doc.activeElement).toBe(outside);
    a.focus();
    followFocus(b);
    expect(doc.activeElement).toBe(b);
    followFocus(null);
    expect(doc.activeElement).toBe(b);
  });

  it('takes focus whatever had it', () => {
    const { doc, items, outside } = fakeGroup();
    outside.focus();
    takeFocus(items[1] as unknown as HTMLElement);
    expect(doc.activeElement).toBe(items[1]);
  });

  it('makes only the current item a Tab stop with the focus mark', () => {
    expect(rovingItem(true)).toEqual({ tabIndex: 0, 'data-focused': '', ref: followFocus });
    expect(rovingItem(false)).toEqual({ tabIndex: -1 });
    expect(rovingItem(true, takeFocus).ref).toBe(takeFocus);
  });
});
