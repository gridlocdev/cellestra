import { describe, expect, it } from 'vitest';
import { Axis } from '../src/grid/axis.ts';
import {
  firstVisible,
  gridKeyAction,
  growOverMerges,
  stepFrom,
  type StepSheet,
} from '../src/grid/navigation.ts';

/** Ten rows and columns, with the given ones hidden (zero size). */
function axes(hiddenRows: number[] = [], hiddenColumns: number[] = []) {
  return {
    rows: new Axis(
      10,
      20,
      hiddenRows.map((i) => [i, 0] as [number, number]),
    ),
    columns: new Axis(
      10,
      64,
      hiddenColumns.map((i) => [i, 0] as [number, number]),
    ),
  };
}

const plain: StepSheet = { stepOverMerges: (r, c, dr, dc) => [r + dr, c + dc] };

const key = (
  k: string,
  mods: Partial<{ shiftKey: boolean; altKey: boolean; metaKey: boolean; ctrlKey: boolean }> = {},
) => ({
  key: k,
  shiftKey: false,
  altKey: false,
  metaKey: false,
  ctrlKey: false,
  ...mods,
});

describe('stepping the active cell', () => {
  it('steps over hidden rows and columns', () => {
    expect(stepFrom({ row: 1, column: 1 }, 1, 0, 'cell', axes([2, 3]), plain)).toEqual({ row: 4, column: 1 });
    expect(stepFrom({ row: 1, column: 1 }, 0, 1, 'cell', axes([], [2]), plain)).toEqual({
      row: 1,
      column: 3,
    });
  });

  it('stays put rather than land on a hidden line at the edge', () => {
    expect(stepFrom({ row: 8, column: 0 }, 1, 0, 'cell', axes([9]), plain)).toEqual({ row: 8, column: 0 });
  });

  it('moves out of a merge in one step', () => {
    const merged: StepSheet = { stepOverMerges: (r, c, dr) => (r === 0 && dr === 1 ? [3, c] : [r + dr, c]) };
    expect(stepFrom({ row: 0, column: 0 }, 1, 0, 'cell', axes(), merged)).toEqual({ row: 3, column: 0 });
  });

  it('counts a page in visible rows', () => {
    expect(stepFrom({ row: 0, column: 2 }, 3, 0, 'page', axes([1]), plain)).toEqual({ row: 4, column: 2 });
  });

  it('extends by visible lines, ignoring merges', () => {
    const merged: StepSheet = { stepOverMerges: () => [9, 9] };
    expect(stepFrom({ row: 0, column: 0 }, 1, 1, 'extend', axes([1]), merged)).toEqual({ row: 2, column: 1 });
  });

  it('finds the first visible line', () => {
    expect(firstVisible(axes([0, 1]).rows)).toBe(2);
    expect(firstVisible(axes().columns)).toBe(0);
  });

  it('grows a range over merges that straddle it, through chains', () => {
    const merges = [
      { range: { top: 0, left: 1, bottom: 1, right: 2 } },
      { range: { top: 1, left: 2, bottom: 3, right: 3 } },
    ];
    expect(growOverMerges(merges, { top: 0, left: 0, bottom: 0, right: 1 })).toEqual({
      top: 0,
      left: 0,
      bottom: 3,
      right: 3,
    });
  });
});

describe('keys on the grid', () => {
  it('moves with arrows, stretches with Shift, jumps with ⌘', () => {
    expect(gridKeyAction(key('ArrowDown'))).toMatchObject({
      kind: 'move',
      rows: 1,
      columns: 0,
      extend: false,
    });
    expect(gridKeyAction(key('ArrowLeft', { shiftKey: true }))).toMatchObject({ columns: -1, extend: true });
    expect(gridKeyAction(key('ArrowRight', { metaKey: true }))).toMatchObject({ jump: true });
  });

  it('cycles with Enter and Tab, backwards with Shift', () => {
    expect(gridKeyAction(key('Enter', { shiftKey: true }))).toMatchObject({ rows: -1, cycle: true });
    expect(gridKeyAction(key('Tab'))).toMatchObject({ columns: 1, cycle: true });
  });

  it('types printable characters, AltGr ones included, but not shortcuts', () => {
    expect(gridKeyAction(key('a'))).toEqual({ kind: 'type', text: 'a' });
    expect(gridKeyAction(key('@', { ctrlKey: true, altKey: true }))).toEqual({ kind: 'type', text: '@' });
    expect(gridKeyAction(key('c', { metaKey: true }))).toBeNull();
    expect(gridKeyAction(key('c', { ctrlKey: true }))).toBeNull();
    expect(gridKeyAction(key('Shift'))).toBeNull();
  });

  it('maps editing, clearing, paging and Home', () => {
    expect(gridKeyAction(key('F2'))).toEqual({ kind: 'edit' });
    expect(gridKeyAction(key('Backspace'))).toEqual({ kind: 'clear' });
    expect(gridKeyAction(key('PageUp', { shiftKey: true }))).toEqual({
      kind: 'page',
      down: false,
      extend: true,
    });
    expect(gridKeyAction(key('Home', { ctrlKey: true }))).toEqual({ kind: 'home', top: true, extend: false });
    expect(gridKeyAction(key('Escape'))).toEqual({ kind: 'escape' });
  });
});
