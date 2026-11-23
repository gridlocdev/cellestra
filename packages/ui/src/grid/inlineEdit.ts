import type { FunctionMetadata, Worksheet } from '@cellestra/core';
import {
  columnCompletion,
  formulaReferences,
  functionPrefix,
  functionSuggestions,
  isReferenceExpected,
  referenceText,
  signatureHint,
  type SignaturePart,
} from './formulaEdit.ts';
import { rangeOf, type GridRange } from '../session/selection.ts';

/**
 * The grid's inline editor, without the DOM: which cell is being edited, committing or cancelling,
 * the editor's keys, pointing at cells while typing a formula, and the formula aids (reference
 * outlines, function suggestions, the argument hint). The grid owns the textarea and the screen;
 * it hands the textarea in as `EditorText` and is told through `InlineEditView` what changed.
 */

export interface CellAt {
  readonly row: number;
  readonly column: number;
}

/** What the model needs of the editor's text field (a textarea has all of it). */
export interface EditorText {
  value: string;
  readonly selectionStart: number;
  readonly selectionEnd: number;
  setRangeText(text: string, start: number, end: number, mode: 'end'): void;
  setSelectionRange(start: number, end: number): void;
}

/** The formula aids under the editor: function names matching what is typed, or the argument hint. */
export interface FormulaAssist {
  readonly suggestions: readonly FunctionMetadata[];
  readonly index: number;
  readonly hint: readonly SignaturePart[] | null;
}

/** Pointing while typing a formula: where the reference went in the text, and what it covers. */
export interface Point {
  readonly anchor: CellAt;
  readonly cursor: CellAt;
  readonly range: GridRange;
  readonly start: number;
  readonly length: number;
}

/** What the edit reads from the document. */
export interface InlineEditHost {
  readonly sheet: Worksheet;
  /** What the editor starts with (`=SUM(A1:A3)` for formulas, the raw input otherwise). */
  editText(row: number, column: number): string;
  /** Return false to reject the entry (validation): the editor stays open. */
  commit(row: number, column: number, text: string): boolean | void;
  /** The editor's text as it is typed, and null once the edit ends. */
  draft?(text: string | null): void;
  /** The sheets a 3D range from `first` to `last` spans, in tab order; null when either is missing. */
  sheetsBetween?(first: string, last: string): readonly string[] | null;
}

/** What the grid does on screen when the edit changes. */
export interface InlineEditView {
  editingChanged(at: CellAt | null): void;
  assistChanged(assist: FormulaAssist | null): void;
  /** The editor's text changed: fit the editor to it. */
  refit(): void;
  /** Reference outlines or the pointed range changed. */
  redraw(): void;
  reveal(row: number, column: number): void;
}

/** Moving one visible line from another, as the viewport does (skipping hidden lines). */
export interface LineStepper {
  readonly rows: { step(from: number, by: number): number };
  readonly columns: { step(from: number, by: number): number };
}

export interface EditorKey {
  readonly key: string;
  readonly shiftKey: boolean;
  readonly altKey: boolean;
  readonly metaKey: boolean;
  readonly ctrlKey: boolean;
}

/**
 * What an editor key does: `handled` (the model did it; prevent the default), `end` (commit or
 * cancel, then move the active cell by `move`), `newline` (insert a line break), or null (let the
 * text field have it).
 */
export type EditorKeyResult =
  | { readonly kind: 'handled' }
  | {
      readonly kind: 'end';
      readonly commit: boolean;
      readonly move: readonly [number, number];
      /** Enter and Tab: the step cycles in a selected range and follows a Tab run; arrows just step. */
      readonly cycle: boolean;
    }
  | { readonly kind: 'newline' }
  | null;

/** How an `end` went: there was no edit, it closed, or the host refused the value. */
export type EndResult = 'none' | 'closed' | 'refused';

const ARROWS: Record<string, readonly [number, number]> = {
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1],
};

export interface InlineEditDeps {
  readonly host: InlineEditHost;
  readonly editor: EditorText | null;
  readonly view: InlineEditView;
}

export class InlineEdit {
  /** The cell being edited, or null. The source of truth for handlers: blur can fire after a commit. */
  at: CellAt | null = null;
  /** The editor was opened by a dead key or IME and has not received a character yet. */
  composing = false;
  /** The reference the last click or arrow put in the formula. */
  point: Point | null = null;
  /** The references of the formula being typed, outlined in colour. */
  references: { range: GridRange; color: string }[] = [];
  assist: FormulaAssist | null = null;
  private deps: InlineEditDeps;

  constructor(deps: InlineEditDeps) {
    this.deps = deps;
  }

  /** Swap in the grid's current host, editor and view (they change between renders). */
  connect(deps: InlineEditDeps): void {
    this.deps = deps;
  }

  private get editor(): EditorText | null {
    return this.deps.editor;
  }

  /**
   * Open the editor on `at` with `initial` (or the cell's own text). `show` puts the editor on
   * screen and focuses it; the caret then goes to the end and the aids are worked out.
   */
  begin(at: CellAt, initial: string | undefined, show: () => void): void {
    const editor = this.editor!;
    const { host } = this.deps;
    editor.value = initial ?? host.editText(at.row, at.column);
    this.setAt(at);
    host.draft?.(editor.value);
    show();
    const end = editor.value.length;
    editor.setSelectionRange(end, end);
    this.refreshAids();
  }

  /** Commit (or cancel) the open edit. */
  end(commit: boolean): EndResult {
    const at = this.at;
    if (!at) return 'none';
    const text = this.editor!.value;
    // An edit a dead key or IME opened that never produced a character leaves the cell as it was.
    if (this.composing && text === '') commit = false;
    this.composing = false;
    if (commit && this.deps.host.commit(at.row, at.column, text) === false) return 'refused';
    this.setAt(null);
    this.deps.host.draft?.(null);
    this.point = null;
    this.references = [];
    this.setAssist(null);
    return 'closed';
  }

  /** The editor's text changed by typing (`typed`) or otherwise (paste, delete, IME). */
  input(typed: boolean): void {
    const editor = this.editor!;
    // Typing ends pointing: the next click or arrow inserts a new reference.
    this.point = null;
    if (editor.value !== '') this.composing = false;
    // Type-ahead: appending to plain text offers the nearest entry above that starts with it,
    // selected, so the next keystroke replaces it and Enter keeps it.
    const at = this.at;
    const text = editor.value;
    if (typed && at && !text.startsWith('=') && editor.selectionStart === text.length) {
      const match = columnCompletion(this.deps.host.sheet, at.row, at.column, text);
      if (match) {
        editor.value = match;
        editor.setSelectionRange(text.length, match.length);
      }
    }
    this.deps.host.draft?.(editor.value);
    this.deps.view.refit();
    this.refreshAids();
  }

  /** Reference outlines, function suggestions and the argument hint, for the text and caret. */
  refreshAids(): void {
    const editor = this.editor;
    if (!editor || !this.at) return;
    const text = editor.value;
    const caret = editor.selectionStart;
    const formula = text.startsWith('=');
    this.references = formula
      ? formulaReferences(text, this.deps.host.sheet.name, (first, last) =>
          this.deps.host.sheetsBetween ? this.deps.host.sheetsBetween(first, last) : null,
        )
      : [];
    const prefix = formula && editor.selectionEnd === caret ? functionPrefix(text, caret) : null;
    const suggestions = prefix ? functionSuggestions(prefix.prefix) : [];
    const hint = formula ? (signatureHint(text, caret)?.parts ?? null) : null;
    const prev = this.assist;
    if (suggestions.length === 0 && !hint) this.setAssist(null);
    else
      this.setAssist({
        suggestions,
        // Keep the highlighted name while it still matches.
        index: Math.max(
          0,
          suggestions.findIndex((m) => m.name === prev?.suggestions[prev.index]?.name),
        ),
        hint,
      });
    this.deps.view.redraw();
  }

  /** Put the highlighted function name (and its opening parenthesis) in place of what was typed. */
  acceptSuggestion(index = this.assist?.index ?? 0): void {
    const editor = this.editor!;
    const meta = this.assist?.suggestions[index];
    const prefix = functionPrefix(editor.value, editor.selectionStart);
    if (!meta || !prefix) return;
    editor.setRangeText(`${meta.name}(`, prefix.start, editor.selectionStart, 'end');
    this.deps.host.draft?.(editor.value);
    this.deps.view.refit();
    this.point = null;
    this.refreshAids();
  }

  /** Whether the caret sits right after the reference the last click or arrow put in. */
  isPointing(): boolean {
    const p = this.point;
    const editor = this.editor;
    return (
      !!p &&
      !!editor &&
      editor.selectionStart === p.start + p.length &&
      editor.selectionEnd === editor.selectionStart
    );
  }

  /** Whether a click or an arrow should put a cell reference in the formula being typed. */
  canPoint(): boolean {
    const editor = this.editor;
    if (!editor || !this.at || !editor.value.startsWith('=')) return false;
    if (editor.selectionStart !== editor.selectionEnd) return false;
    return this.isPointing() || isReferenceExpected(editor.value, editor.selectionStart);
  }

  /** The point to stretch from (⇧-click, ⇧-arrow) while still pointing; else null. */
  get activePoint(): Point | null {
    return this.isPointing() ? this.point : null;
  }

  /**
   * Point at a range: insert its reference at the caret, or replace the one just pointed at, so
   * consecutive pokes move the reference rather than piling them up.
   */
  pointAt(anchor: CellAt, to: CellAt): void {
    const editor = this.editor!;
    const range = rangeOf(anchor.row, anchor.column, to.row, to.column);
    const text = referenceText(range);
    const p = this.activePoint;
    const start = p ? p.start : editor.selectionStart;
    const end = p ? p.start + p.length : editor.selectionEnd;
    editor.setRangeText(text, start, end, 'end');
    this.point = { anchor, cursor: to, range, start, length: text.length };
    this.deps.host.draft?.(editor.value);
    this.deps.view.refit();
    this.deps.view.reveal(to.row, to.column);
    this.refreshAids();
  }

  /** A key in the editor, after the shortcuts had their turn. */
  key(e: EditorKey, lines: () => LineStepper): EditorKeyResult {
    const plain = !e.altKey && !e.metaKey && !e.ctrlKey;
    // The function-name list takes ↑/↓ to choose, Tab to accept, and Escape to close.
    const assist = this.assist;
    const list = assist?.suggestions ?? [];
    if (assist && list.length > 0 && plain) {
      const at = assist.index;
      if ((e.key === 'ArrowDown' || e.key === 'ArrowUp') && !e.shiftKey) {
        const step = e.key === 'ArrowDown' ? 1 : -1;
        this.setAssist({ ...assist, index: (at + step + list.length) % list.length });
        return { kind: 'handled' };
      }
      if (e.key === 'Tab' && !e.shiftKey) {
        this.acceptSuggestion(at);
        return { kind: 'handled' };
      }
      if (e.key === 'Escape') {
        this.setAssist(assist.hint ? { suggestions: [], index: 0, hint: assist.hint } : null);
        return { kind: 'handled' };
      }
    }
    // Pointing: in a formula where a reference can go, the arrows move a reference instead of
    // the caret (⇧ stretches it to a range).
    const arrow = ARROWS[e.key];
    if (arrow && plain && this.canPoint()) {
      const p = this.activePoint;
      const from = p ? p.cursor : this.at!;
      const v = lines();
      const to = {
        row: arrow[0] ? v.rows.step(from.row, arrow[0]) : from.row,
        column: arrow[1] ? v.columns.step(from.column, arrow[1]) : from.column,
      };
      this.pointAt(e.shiftKey && p ? p.anchor : to, to);
      return { kind: 'handled' };
    }
    if (e.key === 'Escape') return { kind: 'end', commit: false, move: [0, 0], cycle: false };
    if (e.key === 'Enter' && !e.shiftKey && !e.altKey)
      return { kind: 'end', commit: true, move: [1, 0], cycle: true };
    // ⌥↩ / ⇧↩ start a new line in the cell.
    if (e.key === 'Enter') return { kind: 'newline' };
    if (e.key === 'Tab') return { kind: 'end', commit: true, move: [0, e.shiftKey ? -1 : 1], cycle: true };
    if ((e.key === 'ArrowUp' || e.key === 'ArrowDown') && !e.shiftKey && plain) {
      // ↑/↓ commit and move a cell, like Enter. In multi-line text they move between lines
      // first, and only leave the cell from the first or last line.
      const editor = this.editor!;
      const up = e.key === 'ArrowUp';
      const caret = editor.selectionStart;
      const onEdgeLine = up
        ? !editor.value.slice(0, caret).includes('\n')
        : !editor.value.slice(caret).includes('\n');
      return onEdgeLine ? { kind: 'end', commit: true, move: [up ? -1 : 1, 0], cycle: false } : null;
    }
    return null;
  }

  private setAt(at: CellAt | null): void {
    this.at = at;
    this.deps.view.editingChanged(at);
  }

  private setAssist(assist: FormulaAssist | null): void {
    this.assist = assist;
    this.deps.view.assistChanged(assist);
  }
}
