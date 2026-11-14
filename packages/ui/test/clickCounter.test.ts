import { describe, expect, it } from 'vitest';
import { ClickCounter } from '../src/grid/clickCounter.ts';

describe('ClickCounter', () => {
  it('counts a quick second press in the same spot as a double click', () => {
    const clicks = new ClickCounter();
    expect(clicks.press(1000, 100, 20, 0)).toBe(1);
    expect(clicks.press(1200, 101, 21, 0)).toBe(2);
    expect(clicks.press(1350, 101, 21, 0)).toBe(3);
  });

  it('starts over after a pause or a move', () => {
    const clicks = new ClickCounter();
    clicks.press(1000, 100, 20, 0);
    expect(clicks.press(1600, 100, 20, 0)).toBe(1);
    expect(clicks.press(1700, 120, 20, 0)).toBe(1);
  });

  it("trusts the browser's count when it gives one", () => {
    const clicks = new ClickCounter();
    expect(clicks.press(1000, 100, 20, 2)).toBe(2);
    expect(clicks.press(5000, 300, 20, 1)).toBe(1);
  });
});
