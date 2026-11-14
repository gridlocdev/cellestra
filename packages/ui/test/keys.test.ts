import { describe, expect, it } from 'vitest';
import { startsComposition } from '../src/grid/keys.ts';

const key = (key: string, keyCode = 0, isComposing = false) => ({ key, keyCode, isComposing });

describe('startsComposition', () => {
  it('is true for dead keys and IME keys', () => {
    expect(startsComposition(key('Dead', 222))).toBe(true);
    expect(startsComposition(key('Process', 229))).toBe(true);
    expect(startsComposition(key('a', 229))).toBe(true);
    expect(startsComposition(key('n', 78, true))).toBe(true);
  });

  it('is false for plain characters and commands', () => {
    expect(startsComposition(key('a', 65))).toBe(false);
    expect(startsComposition(key('Enter', 13))).toBe(false);
    expect(startsComposition(key('ArrowDown', 40))).toBe(false);
  });
});
