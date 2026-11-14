import { describe, expect, it } from 'vitest';
import { trapIndex } from '../src/components/focusTrap.ts';

describe('dialog focus trap', () => {
  it('moves to the next stop and wraps from the last to the first', () => {
    expect(trapIndex(3, 0, false)).toBe(1);
    expect(trapIndex(3, 2, false)).toBe(0);
  });

  it('moves back with Shift+Tab and wraps from the first to the last', () => {
    expect(trapIndex(3, 1, true)).toBe(0);
    expect(trapIndex(3, 0, true)).toBe(2);
  });

  it('enters at an end when focus is on none of the stops', () => {
    expect(trapIndex(3, -1, false)).toBe(0);
    expect(trapIndex(3, -1, true)).toBe(2);
  });

  it('has nowhere to go without stops', () => {
    expect(trapIndex(0, -1, false)).toBeNull();
  });
});
