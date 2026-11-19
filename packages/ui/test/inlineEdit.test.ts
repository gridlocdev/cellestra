import { describe, expect, it, vi } from 'vitest';
import { CellValue, Workbook } from '@cellestra/core';
import {
  InlineEdit,
  type CellAt,
  type EditorKey,
  type EditorText,
  type FormulaAssist,
  type LineStepper,
} from '../src/grid/inlineEdit.ts';

/** A text field with a caret, like the textarea the grid hands in. */
class FakeEditor implements EditorText {
  value = '';
  selectionStart = 0;
  selectionEnd = 0;
  setRangeText(text: string, start: number, end: number): void {
    this.value = this.value.slice(0, start) + text + this.value.slice(end);
    this.selectionStart = this.selectionEnd = start + text.length;
  }
  setSelectionRange(start: number, end: number): void {
    this.selectionStart = start;
    this.selectionEnd = end;
  }
  /** Type `text` at the caret. */
  type(text: string): void {
    this.setRangeText(text, this.selectionStart, this.selectionEnd);
  }
}

const lines: LineStepper = {
  rows: { step: (from, by) => Math.max(0, from + by) },
  columns: { step: (from, by) => Math.max(0, from + by) },
};

const key = (k: string, mods: Partial<EditorKey> = {}): EditorKey => ({
  key: k,
  shiftKey: false,
  altKey: false,
  metaKey: false,
  ctrlKey: false,
  ...mods,
});

function setup(commit: (row: number, column: number, text: string) => boolean | void = () => {}) {
  const sheet = Workbook.createNew().activeSheet;
  const editor = new FakeEditor();
  const commits: string[] = [];
  const drafts: (string | null)[] = [];
  const view = {
    at: null as CellAt | null,
    assist: null as FormulaAssist | null,
    editingChanged: (at: CellAt | null) => (view.at = at),
    assistChanged: (a: FormulaAssist | null) => (view.assist = a),
    refit: vi.fn(),
    redraw: vi.fn(),
    reveal: vi.fn(),
  };
  const edit = new InlineEdit({
    host: {
      sheet,
      editText: (row, column) => (row === 0 && column === 0 ? 'start' : ''),
      commit: (row, column, text) => {
        commits.push(text);
        return commit(row, column, text);
      },
      draft: (text) => drafts.push(text),
    },
    editor,
    view,
  });
  const show = vi.fn();
  return { sheet, editor, edit, view, commits, drafts, show };
}

describe('beginning and ending an edit', () => {
  it("opens with the cell's text and the caret at the end, and commits it on Enter", () => {
    const { edit, editor, view, commits, drafts, show } = setup();
    edit.begin({ row: 0, column: 0 }, undefined, show);
    expect(show).toHaveBeenCalledOnce();
    expect(editor.value).toBe('start');
    expect(editor.selectionStart).toBe(5);
    expect(view.at).toEqual({ row: 0, column: 0 });
    editor.type('ed');
    expect(edit.key(key('Enter'), () => lines)).toEqual({
      kind: 'end',
      commit: true,
      move: [1, 0],
      cycle: true,
    });
    expect(edit.end(true)).toBe('closed');
    expect(commits).toEqual(['started']);
    expect(view.at).toBeNull();
    expect(drafts).toEqual(['start', null]);
  });

  it('keeps the editor open when the host refuses the value', () => {
    const { edit } = setup(() => false);
    edit.begin({ row: 1, column: 1 }, 'bad', () => {});
    expect(edit.end(true)).toBe('refused');
    expect(edit.at).toEqual({ row: 1, column: 1 });
  });

  it('leaves the cell alone when a dead key opened the editor and nothing was typed', () => {
    const { edit, commits } = setup();
    edit.begin({ row: 1, column: 1 }, '', () => {});
    edit.composing = true;
    expect(edit.end(true)).toBe('closed');
    expect(commits).toEqual([]);
  });

  it('reports no edit to end when none is open', () => {
    expect(setup().edit.end(true)).toBe('none');
  });
});

describe('editor keys', () => {
  it('Escape cancels, Tab commits sideways, ⌥↩ breaks the line', () => {
    const { edit } = setup();
    edit.begin({ row: 3, column: 3 }, 'x', () => {});
    expect(edit.key(key('Escape'), () => lines)).toEqual({
      kind: 'end',
      commit: false,
      move: [0, 0],
      cycle: false,
    });
    expect(edit.key(key('Tab', { shiftKey: true }), () => lines)).toEqual({
      kind: 'end',
      commit: true,
      move: [0, -1],
      cycle: true,
    });
    expect(edit.key(key('Enter', { altKey: true }), () => lines)).toEqual({ kind: 'newline' });
  });

  it('moves between lines of multi-line text before ↓ leaves the cell', () => {
    const { edit, editor } = setup();
    edit.begin({ row: 3, column: 3 }, 'one\ntwo', () => {});
    editor.setSelectionRange(1, 1);
    expect(edit.key(key('ArrowDown'), () => lines)).toBeNull();
    editor.setSelectionRange(5, 5);
    expect(edit.key(key('ArrowDown'), () => lines)).toEqual({
      kind: 'end',
      commit: true,
      move: [1, 0],
      cycle: false,
    });
  });
});

describe('pointing while typing a formula', () => {
  it('puts a reference in with an arrow, moves it with the next, and stretches it with ⇧', () => {
    const { edit, editor, view } = setup();
    edit.begin({ row: 4, column: 0 }, '=SUM(', () => {});
    expect(edit.key(key('ArrowUp'), () => lines)).toEqual({ kind: 'handled' });
    expect(editor.value).toBe('=SUM(A4');
    edit.key(key('ArrowUp'), () => lines);
    expect(editor.value).toBe('=SUM(A3');
    edit.key(key('ArrowUp', { shiftKey: true }), () => lines);
    expect(editor.value).toBe('=SUM(A2:A3');
    expect(edit.point?.range).toEqual({ top: 1, left: 0, bottom: 2, right: 0 });
    expect(view.reveal).toHaveBeenLastCalledWith(1, 0);
  });

  it('stops pointing once something is typed, so the next arrow inserts a new reference', () => {
    const { edit, editor } = setup();
    edit.begin({ row: 0, column: 1 }, '=', () => {});
    edit.pointAt({ row: 0, column: 0 }, { row: 0, column: 0 });
    expect(edit.isPointing()).toBe(true);
    editor.type('+');
    edit.input(true);
    expect(edit.point).toBeNull();
    expect(edit.canPoint()).toBe(true);
    edit.pointAt({ row: 1, column: 0 }, { row: 1, column: 0 });
    expect(editor.value).toBe('=A1+A2');
  });

  it('outlines the references of the formula', () => {
    const { edit } = setup();
    edit.begin({ row: 5, column: 5 }, '=A1+B2:C3', () => {});
    expect(edit.references.map((r) => r.range)).toEqual([
      { top: 0, left: 0, bottom: 0, right: 0 },
      { top: 1, left: 1, bottom: 2, right: 2 },
    ]);
  });
});

describe('formula assist', () => {
  it('suggests function names, moves through them, and Tab inserts the chosen one', () => {
    const { edit, editor, view } = setup();
    edit.begin({ row: 0, column: 1 }, '=SU', () => {});
    const names = view.assist!.suggestions.map((m) => m.name);
    expect(names).toContain('SUM');
    expect(edit.key(key('ArrowDown'), () => lines)).toEqual({ kind: 'handled' });
    expect(view.assist!.index).toBe(1);
    edit.key(key('Tab'), () => lines);
    expect(editor.value).toBe(`=${names[1]}(`);
  });

  it('closes the list on Escape and keeps the argument hint', () => {
    const { edit, editor, view } = setup();
    edit.begin({ row: 0, column: 1 }, '=SUM(1, SU', () => {});
    expect(view.assist!.suggestions.length).toBeGreaterThan(0);
    expect(edit.key(key('Escape'), () => lines)).toEqual({ kind: 'handled' });
    expect(view.assist!.suggestions).toEqual([]);
    expect(view.assist!.hint).not.toBeNull();
    expect(editor.value).toBe('=SUM(1, SU');
  });
});

describe('type-ahead', () => {
  it('offers the entry above that starts with what is typed, selected', () => {
    const { edit, editor, sheet } = setup();
    sheet.setValue(0, 2, CellValue.text('apple'));
    edit.begin({ row: 1, column: 2 }, '', () => {});
    editor.type('ap');
    edit.input(true);
    expect(editor.value).toBe('apple');
    expect([editor.selectionStart, editor.selectionEnd]).toEqual([2, 5]);
  });
});
