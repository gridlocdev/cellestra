import { useEffect, useLayoutEffect, useRef, useState } from 'preact/hooks';
import { effect, type ReadonlySignal } from '@preact/signals';
import {
  Defaults,
  type CellStyle,
  type CellValue,
  type FunctionMetadata,
  type StylePool,
  type Worksheet,
} from '@cellestra/core';
import { Axis } from './axis.ts';
import { chromeButtons, hitChrome, type ChromeButton } from './chrome.ts';
import {
  columnCompletion,
  fillDownTarget,
  fillTarget,
  formulaReferences,
  functionPrefix,
  functionSuggestions,
  isReferenceExpected,
  referenceText,
  signatureHint,
  type SignaturePart,
} from './formulaEdit.ts';
import {
  checkboxRect,
  fillHandleRect,
  renderGrid,
  wrap,
  CELL_PAD_X,
  CELL_PAD_Y,
  MARQUEE_DASH,
  type DisplayText,
  type GridOverlay,
} from './render.ts';
import { ContentHeights } from './rowHeights.ts';
import { insertLineBreak } from '../components/lineBreak.ts';
import { Viewport, HEADER_HEIGHT, HEADER_WIDTH, type Rect } from './viewport.ts';
import type { GridPalette } from './palette.ts';
import { MAX_COLUMN, MAX_ROW, rangeOf, type GridRange, type Selection } from '../session/selection.ts';

/** What the grid needs from the document. The grid never edits the model itself. */
export interface GridHost {
  readonly sheet: ReadonlySignal<Worksheet>;
  readonly styles: StylePool;
  /** Bumped on every model change; the grid redraws when it moves. */
  readonly version: ReadonlySignal<number>;
  /**
   * The cells of `sheet` changed since `version` (edited or recalculated), or null when that is
   * unknown. Lets the grid re-measure row heights for just those cells.
   */
  cellsChangedSince?(sheet: string, version: number): readonly { row: number; column: number }[] | null;
  display(value: CellValue, style: CellStyle): DisplayText;
  /**
   * The style a cell shows once its conditional formats are applied (`style` when none holds).
   * Only called for styles that carry conditional maps.
   */
  conditionalStyle?(style: CellStyle, row: number, column: number): CellStyle;
  /** What the inline editor starts with (`=SUM(A1:A3)` for formulas, the raw input otherwise). */
  editText(row: number, column: number): string;
  /**
   * The inline editor's text as it is typed (so the formula bar can mirror it), and null once
   * the edit ends.
   */
  draft?(text: string | null): void;
  /** Return false to reject the entry (validation): the editor stays open. */
  commit(row: number, column: number, text: string): boolean | void;
  clearSelection(): void;
  /** Edge navigation (⌘-arrow): the cell to jump to. */
  edge(row: number, column: number, rowStep: number, columnStep: number): { row: number; column: number };
  /** Key not handled by the grid (shortcuts) → command catalog. Return true when handled. */
  shortcut(e: KeyboardEvent, inText: boolean): boolean;
  /** Column/row resize from the header borders. */
  resizeColumn(column: number, width: number): void;
  resizeRow(row: number, height: number): void;
  /** A filter chevron or validation arrow was clicked; `anchor` is its rect in client px. */
  chrome?(button: ChromeButton, anchor: { x: number; y: number; w: number; h: number }): void;
  /** ⌘/Ctrl-click on a cell with a hyperlink. Return true when it opened something. */
  openLink?(row: number, column: number): boolean;
  /** Whether a cell has a hyperlink (the pointer cursor while ⌘/Ctrl is held). */
  hasLink?(row: number, column: number): boolean;
  /**
   * Tick or clear checkboxes: the one at `at` (a click on its box), or every checkbox in the
   * selection (Space). Return false when there was none, so the key does what it otherwise would.
   */
  toggleCheckboxes?(at?: { row: number; column: number }): boolean;
  /** A fill-handle drag or double-click: copy `source` across `target`, then select the target. */
  fill?(source: GridRange, target: GridRange): void;
  /** The copy/cut source to outline with marching ants, until paste, Escape, or an edit. */
  readonly marquee?: ReadonlySignal<{ sheet: string; range: GridRange; cut: boolean } | null>;
  /** Escape or a new edit: drop the marquee (and with it a pending cut). */
  clearMarquee?(): void;
  /** A style shown over the selection without applying it (hovering a font size or colour). */
  readonly stylePreview?: ReadonlySignal<((style: CellStyle) => CellStyle) | null>;
  /** Double-click on a column header's right edge: fit the column to its content. */
  autofitColumn?(column: number): void;
  /** Right-click (or Ctrl-click on macOS); the selection already covers what was clicked. */
  contextMenu?(zone: 'cell' | 'columnHeader' | 'rowHeader', at: { x: number; y: number }): void;
  /** The hover tooltip for a cell (an error's reason, a link's hint), or null. */
  hoverTip?(row: number, column: number): string | null;
}

export interface GridProps {
  readonly host: GridHost;
  readonly selection: Selection;
  readonly zoom: ReadonlySignal<number>;
  readonly palette: ReadonlySignal<GridPalette>;
  readonly showFormulas?: ReadonlySignal<boolean>;
  /** Imperative handle: the formula bar asks the grid to start / end editing. */
  readonly controller?: { current: GridController | null };
}

export interface GridController {
  focus(): void;
  beginEdit(initial?: string): void;
  isEditing(): boolean;
  /**
   * Commit the open inline edit, if any, before something else changes the sheet or document.
   * False when the value was refused (a stop validation rule) and the editor stays open.
   */
  commitEdit(): boolean;
}

const MIN_COLUMN_WIDTH = 16;
const MIN_ROW_HEIGHT = 8;
/** Must match `.cs-grid-editor` in styles.css. */
const EDITOR_BORDER = 2;

const EDITOR_LINE_HEIGHT = 1.2;
const RESIZE_GRAB = 4;
/** Extra px around the 7px fill handle that still grabs it. */
const HANDLE_GRAB = 3;
/** How long the pointer rests on a cell before its tooltip shows. */
const TIP_DELAY_MS = 450;
const MARQUEE_STEP_MS = 90;
const IS_MAC = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform);

interface Axes {
  readonly columns: Axis;
  readonly rows: Axis;
  /** The last row and column with non-default layout (−1 for none), for the scroll extent. */
  readonly lastRow: number;
  readonly lastColumn: number;
}

function axesFor(sheet: Worksheet, content: ReadonlyMap<number, number>): Axes {
  const cols: [number, number][] = [];
  let lastColumn = -1;
  for (const [i, c] of sheet.columns) {
    cols.push([i, c.hidden ? 0 : c.width]);
    if (!c.isDefault) lastColumn = Math.max(lastColumn, i);
  }
  const rows: [number, number][] = [];
  let lastRow = -1;
  for (const [i, r] of sheet.rows) {
    rows.push([i, r.hidden ? 0 : Math.max(r.height, content.get(i) ?? 0)]);
    if (!r.isDefault) lastRow = Math.max(lastRow, i);
  }
  // Rows with no layout of their own that their content makes taller.
  for (const [i, h] of content) if (!sheet.rows.has(i)) rows.push([i, h]);
  return {
    columns: new Axis(MAX_COLUMN + 1, 64, cols),
    rows: new Axis(MAX_ROW + 1, 20, rows),
    lastRow,
    lastColumn,
  };
}

/** The bottom-right corner the scroller must reach: cells, merges and laid-out lines. */
function contentEnd(sheet: Worksheet, axes: Axes): { bottom: number; right: number } {
  const cells = sheet.cellExtent();
  let bottom = Math.max(cells?.lastRow ?? 0, axes.lastRow);
  let right = Math.max(cells?.lastColumn ?? 0, axes.lastColumn);
  for (const m of sheet.merges) {
    bottom = Math.max(bottom, m.range.bottom);
    right = Math.max(right, m.range.right);
  }
  return { bottom, right };
}

export function Grid({ host, selection, zoom, palette, showFormulas, controller }: GridProps) {
  const rootRef = useRef<HTMLDivElement>(null);
  /** The line shown while dragging a column/row edge; the size is committed once, on release. */
  const guideRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const scrollerRef = useRef<HTMLDivElement>(null);
  const spacerRef = useRef<HTMLDivElement>(null);
  const editorRef = useRef<HTMLTextAreaElement>(null);
  const [size, setSize] = useState({ w: 0, h: 0 });
  const [editing, setEditingState] = useState<{ row: number; column: number } | null>(null);
  // The ref is the source of truth for handlers (blur can fire after a commit, before re-render).
  const editingRef = useRef<{ row: number; column: number } | null>(null);
  const setEditing = (v: { row: number; column: number } | null) => {
    editingRef.current = v;
    setEditingState(v);
  };
  const state = useRef({
    scrollX: 0,
    scrollY: 0,
    viewport: null as Viewport | null,
    /** The axes are rebuilt only when row/column layout moves (`layoutStamp`) or the sheet changes. */
    axesStamp: -1,
    axesSheet: null as Worksheet | null,
    axes: null as Axes | null,
    /** The scroll extent for `usedVersion`. */
    usedVersion: -1,
    used: null as { bottom: number; right: number } | null,
    drag: null as null | Drag,
    chrome: [] as ChromeButton[],
    /** Rows their content makes taller, for `contentVersion`. */
    content: new Map<number, number>(),
    /** What `content` is kept up to date by, edit by edit. */
    heights: null as ContentHeights | null,
    contentVersion: -1,
    contentStamp: -1,
    /** A fill-handle drag's target, outlined while dragging. */
    fill: null as GridRange | null,
    /** The references of the formula being typed, outlined in colour. */
    references: [] as { range: GridRange; color: string }[],
    /** Pointing while typing a formula: the reference the last click or arrow put in the text. */
    point: null as Point | null,
    /** Marching-ants dash offset. */
    phase: 0,
    /** The last pointer position over the grid (for the cursor when ⌘ goes down or up). */
    pointer: null as { x: number; y: number } | null,
    tipTimer: 0,
    tipCell: null as { row: number; column: number } | null,
  });
  const [tip, setTip] = useState<{ text: string; x: number; y: number } | null>(null);
  /** The formula aids under the editor: function names matching what is typed, or the argument hint. */
  const [assist, setAssist] = useState<{
    suggestions: readonly FunctionMetadata[];
    index: number;
    hint: readonly SignaturePart[] | null;
  } | null>(null);
  const assistRef = useRef(assist);
  assistRef.current = assist;

  // --- layout -----------------------------------------------------------------------------------
  useLayoutEffect(() => {
    const el = rootRef.current!;
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const measureCtx = useRef<CanvasRenderingContext2D | null>(null);
  const fontCss = (style: CellStyle, zoom = 1) =>
    `${style.font.italic ? 'italic ' : ''}${style.font.bold ? 'bold ' : ''}${((style.font.sizePoints * 96) / 72) * zoom}px "${style.font.family}", "Helvetica Neue", Arial, sans-serif`;

  /** Recompute the rows content makes taller; true when that changed. */
  const refreshContentHeights = (sheet: Worksheet, version: number): boolean => {
    const s = state.current;
    const same = s.heights?.sheet === sheet && s.contentStamp === sheet.layoutStamp;
    if (same && s.contentVersion === version) return false;
    const since = s.contentVersion;
    s.contentVersion = version;
    s.contentStamp = sheet.layoutStamp;
    // Same sheet and layout: re-measure only the cells edited or recalculated since.
    const changed = same ? host.cellsChangedSince?.(sheet.name, since) : null;
    if (changed && s.heights) return s.heights.update(changed);
    measureCtx.current ??= document.createElement('canvas').getContext('2d');
    const ctx = measureCtx.current;
    s.heights = new ContentHeights(sheet, host.styles, {
      display: (value, style) => host.display(value, style).text,
      wrappedLines: (text, style, width) => {
        if (!ctx) return text.split('\n').length;
        ctx.font = fontCss(style);
        return wrap(ctx, text, width).length;
      },
      columnWidth: (c) => sheet.columns.get(c)?.width ?? Defaults.columnWidth,
      textWidth: (text, style) => {
        if (!ctx) return text.length * style.font.sizePoints * 0.6;
        ctx.font = fontCss(style);
        return ctx.measureText(text).width;
      },
    });
    const prev = s.content;
    const next = s.heights.rows;
    s.content = next;
    if (prev.size !== next.size) return true;
    for (const [k, h] of next) if (prev.get(k) !== h) return true;
    return false;
  };

  const viewport = (): Viewport => {
    const s = state.current;
    const version = host.version.value;
    const sheet = host.sheet.value;
    const grew = refreshContentHeights(sheet, version);
    if (grew || !s.axes || s.axesStamp !== sheet.layoutStamp || s.axesSheet !== sheet) {
      s.axes = axesFor(sheet, s.content);
      s.axesStamp = sheet.layoutStamp;
      s.axesSheet = sheet;
      s.usedVersion = -1;
    }
    if (s.usedVersion !== version) {
      s.used = contentEnd(sheet, s.axes);
      s.usedVersion = version;
    }
    return new Viewport({
      columns: s.axes.columns,
      rows: s.axes.rows,
      frozenColumns: sheet.frozenColumns,
      frozenRows: sheet.frozenRows,
      scrollX: s.scrollX,
      scrollY: s.scrollY,
      zoom: zoom.value,
      width: size.w,
      height: size.h,
    });
  };

  const draw = () => {
    const canvas = canvasRef.current;
    if (!canvas || size.w === 0) return;
    const dpr = globalThis.devicePixelRatio || 1;
    if (canvas.width !== Math.round(size.w * dpr) || canvas.height !== Math.round(size.h * dpr)) {
      canvas.width = Math.round(size.w * dpr);
      canvas.height = Math.round(size.h * dpr);
    }
    const ctx = canvas.getContext('2d')!;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    const v = viewport();
    state.current.viewport = v;
    const sel = selection.state.value;
    const chrome = chromeButtons(host.sheet.value, v, sel);
    state.current.chrome = chrome;
    const s = state.current;
    const marquee = host.marquee?.value;
    const preview = host.stylePreview?.value;
    const overlay: GridOverlay = {
      fill: s.fill,
      references: s.references,
      point: s.point?.range ?? null,
      marquee:
        marquee && marquee.sheet === host.sheet.value.name
          ? { range: marquee.range, cut: marquee.cut, phase: s.phase }
          : null,
      stylePreview: preview ? { ranges: sel.ranges, transform: preview } : null,
    };
    renderGrid({
      overlay,
      chrome,
      ctx,
      viewport: v,
      sheet: host.sheet.value,
      styles: host.styles,
      palette: palette.value,
      selection: sel,
      display: (val, style) => host.display(val, style),
      conditional: host.conditionalStyle?.bind(host),
      editing,
      showFormulas: showFormulas?.value ?? false,
    });
    updateSpacer(v);
    placeEditor(v);
  };

  // Spacer: the native scroller's content is the used range plus room to keep scrolling, so the
  // scrollbar thumb stays useful instead of spanning a million empty rows (Calc/Excel behaviour).
  const updateSpacer = (v: Viewport) => {
    const spacer = spacerRef.current;
    if (!spacer) return;
    const used = state.current.used;
    const sel = selection.range.value;
    const z = v.zoom;
    const bottom = Math.max(used?.bottom ?? 0, sel.bottom) + 1;
    const right = Math.max(used?.right ?? 0, sel.right) + 1;
    const h = Math.min(
      v.rows.offset(bottom) * z + v.height * 2 + state.current.scrollY * z,
      v.rows.total * z,
    );
    const w = Math.min(
      v.columns.offset(right) * z + v.width * 2 + state.current.scrollX * z,
      v.columns.total * z,
    );
    spacer.style.height = `${h}px`;
    spacer.style.width = `${w}px`;
  };

  // Redraw on any signal the frame reads.
  useEffect(() => effect(() => draw()), [size, editing, host, selection, zoom, palette, showFormulas]);

  // The marquee's ants march while it is up.
  useEffect(() => {
    const m = host.marquee;
    if (!m) return;
    let timer = 0;
    const stop = effect(() => {
      const on = m.value !== null;
      if (on && !timer)
        timer = window.setInterval(() => {
          state.current.phase = (state.current.phase + 1) % (MARQUEE_DASH[0] + MARQUEE_DASH[1]);
          draw();
        }, MARQUEE_STEP_MS);
      else if (!on && timer) {
        clearInterval(timer);
        timer = 0;
      }
    });
    return () => {
      stop();
      if (timer) clearInterval(timer);
    };
  }, [host, size, editing]);

  // --- scrolling ---------------------------------------------------------------------------------
  // The scroller's position in sheet px is the truth; a scroll event that only echoes a position we
  // set ourselves changes nothing (no flag to get stuck when the browser fires no event).
  const onScroll = () => {
    const el = scrollerRef.current!;
    const x = el.scrollLeft / zoom.value;
    const y = el.scrollTop / zoom.value;
    const s = state.current;
    if (Math.abs(x - s.scrollX) * zoom.value < 0.5 && Math.abs(y - s.scrollY) * zoom.value < 0.5) return;
    s.scrollX = x;
    s.scrollY = y;
    draw();
  };

  // Zoom keeps the same sheet position: the spacer was resized by the redraw, move the scroller to match.
  useEffect(
    () =>
      effect(() => {
        const z = zoom.value;
        const el = scrollerRef.current;
        if (!el) return;
        el.scrollLeft = state.current.scrollX * z;
        el.scrollTop = state.current.scrollY * z;
      }),
    [zoom],
  );

  const reveal = (row: number, column: number) => {
    const v = state.current.viewport ?? viewport();
    const x = v.scrollToReveal('x', column);
    const y = v.scrollToReveal('y', row);
    if (x !== state.current.scrollX || y !== state.current.scrollY) {
      state.current.scrollX = x;
      state.current.scrollY = y;
      const el = scrollerRef.current!;
      draw();
      el.scrollLeft = x * zoom.value;
      el.scrollTop = y * zoom.value;
    }
  };

  // --- editing -----------------------------------------------------------------------------------
  const beginEdit = (initial?: string) => {
    const { row, column } = selection.active.value;
    const editor = editorRef.current!;
    editor.value = initial ?? host.editText(row, column);
    setEditing({ row, column });
    host.clearMarquee?.();
    hideTip();
    host.draft?.(editor.value);
    // Show and focus synchronously, inside the key event: the next keystrokes (and an IME
    // composition) must land in the editor, not the grid.
    reveal(row, column);
    placeEditor(state.current.viewport ?? viewport());
    editor.focus();
    const end = editor.value.length;
    editor.setSelectionRange(end, end);
    refreshFormulaAids();
  };

  /** False when the host rejected the value (a stop validation rule); the editor stays open. */
  const endEdit = (commit: boolean, move: [number, number] = [0, 0], refocus = true): boolean => {
    const at = editingRef.current;
    if (!at) return true;
    const text = editorRef.current!.value;
    if (commit && host.commit(at.row, at.column, text) === false) {
      if (refocus) editorRef.current!.select();
      return false;
    }
    setEditing(null);
    host.draft?.(null);
    state.current.point = null;
    state.current.references = [];
    setAssist(null);
    if (refocus) rootRef.current?.focus();
    if (move[0] || move[1]) moveActive(move[0], move[1], false, true);
    return true;
  };

  const placeEditor = (v: Viewport) => {
    const editor = editorRef.current;
    if (!editor) return;
    const editing = editingRef.current;
    if (!editing) {
      editor.style.display = 'none';
      return;
    }
    const r = host.sheet.value.selectionRangeAt(editing.row, editing.column);
    const rect = v.rangeRect(r.top, r.left, r.bottom, r.right);
    const style = host.styles.get(host.sheet.value.effectiveStyleId(editing.row, editing.column));
    const wraps = style.alignment.wrapText;
    const fontPx = ((style.font.sizePoints * 96) / 72) * v.zoom;
    // The box is exactly the cell, so the vertical padding is whatever room the line leaves
    // (none in a default 20px row), with the text near the bottom as the grid draws it.
    const room = Math.max(0, rect.h - 2 * EDITOR_BORDER - fontPx * EDITOR_LINE_HEIGHT);
    const padBottom = Math.min(room, CELL_PAD_Y * v.zoom);
    Object.assign(editor.style, {
      display: 'block',
      left: `${rect.x}px`,
      top: `${rect.y}px`,
      // Exactly the cell: a textarea's default size (cols × rows) is wider and taller than one.
      width: `${rect.w}px`,
      height: `${rect.h}px`,
      // The 2px border sits inside the cell, so the side padding gives it back: the text then
      // starts, and wraps, exactly where the grid draws it.
      padding: `${room - padBottom}px ${Math.max(0, CELL_PAD_X * v.zoom - EDITOR_BORDER)}px ${padBottom}px`,
      font: `${style.font.italic ? 'italic ' : ''}${style.font.bold ? 'bold ' : ''}${fontPx}px "${style.font.family}", "Helvetica Neue", Arial, sans-serif`,
      // A wrapping cell wraps in the editor too, at the cell's width, as the grid will draw it.
      whiteSpace: wraps ? 'pre-wrap' : 'pre',
      overflowWrap: wraps ? 'anywhere' : 'normal',
    });
    fitEditor(rect, v, wraps);
  };

  /**
   * Grow the editor past its cell when the text needs it, as Calc does: to the right for a long
   * line (up to the grid's right edge) and down for extra lines. Never smaller than the cell.
   * Wrapping text keeps the cell's width and only grows down.
   */
  const fitEditor = (cell: Rect, v: Viewport, wraps: boolean) => {
    const editor = editorRef.current!;
    const wider = editor.scrollWidth - editor.clientWidth;
    if (wider > 0 && !wraps) {
      const room = Math.max(cell.w, v.width - cell.x - 2);
      editor.style.width = `${Math.min(room, cell.w + wider + 2 * CELL_PAD_X * v.zoom)}px`;
    }
    const taller = editor.scrollHeight - editor.clientHeight;
    if (taller > 0) {
      const room = Math.max(cell.h, v.height - cell.y - 2);
      editor.style.height = `${Math.min(room, cell.h + taller)}px`;
    }
  };

  /** Re-fit the editor after its text changed. */
  const refitEditor = () => {
    const v = state.current.viewport;
    if (v && editingRef.current) placeEditor(v);
  };

  // --- formula aids ------------------------------------------------------------------------------
  /** Reference outlines, function suggestions and the argument hint, for the editor's text and caret. */
  const refreshFormulaAids = () => {
    const editor = editorRef.current;
    if (!editor || !editingRef.current) return;
    const text = editor.value;
    const caret = editor.selectionStart;
    const formula = text.startsWith('=');
    state.current.references = formula ? formulaReferences(text, host.sheet.value.name) : [];
    const prefix = formula && editor.selectionEnd === caret ? functionPrefix(text, caret) : null;
    const suggestions = prefix ? functionSuggestions(prefix.prefix) : [];
    const hint = formula ? (signatureHint(text, caret)?.parts ?? null) : null;
    const prev = assistRef.current;
    if (suggestions.length === 0 && !hint) setAssist(null);
    else
      setAssist({
        suggestions,
        // Keep the highlighted name while it still matches.
        index: Math.max(
          0,
          suggestions.findIndex((m) => m.name === prev?.suggestions[prev.index]?.name),
        ),
        hint,
      });
    draw();
  };

  /** Put the highlighted function name (and its opening parenthesis) in place of what was typed. */
  const acceptSuggestion = (index = assistRef.current?.index ?? 0) => {
    const editor = editorRef.current!;
    const meta = assistRef.current?.suggestions[index];
    const prefix = functionPrefix(editor.value, editor.selectionStart);
    if (!meta || !prefix) return;
    const insert = `${meta.name}(`;
    editor.setRangeText(insert, prefix.start, editor.selectionStart, 'end');
    host.draft?.(editor.value);
    refitEditor();
    state.current.point = null;
    refreshFormulaAids();
  };

  /** Whether the caret sits right after the reference the last click or arrow put in. */
  const isPointing = (): boolean => {
    const p = state.current.point;
    const editor = editorRef.current;
    return (
      !!p &&
      !!editor &&
      editor.selectionStart === p.start + p.length &&
      editor.selectionEnd === editor.selectionStart
    );
  };

  /** Whether a click or an arrow should put a cell reference in the formula being typed. */
  const canPoint = (): boolean => {
    const editor = editorRef.current;
    if (!editor || !editingRef.current || !editor.value.startsWith('=')) return false;
    if (editor.selectionStart !== editor.selectionEnd) return false;
    return isPointing() || isReferenceExpected(editor.value, editor.selectionStart);
  };

  /**
   * Point at a range: insert its reference at the caret, or replace the one just pointed at, so
   * consecutive pokes move the reference rather than piling them up.
   */
  const pointAt = (anchor: { row: number; column: number }, to: { row: number; column: number }) => {
    const editor = editorRef.current!;
    const range = rangeOf(anchor.row, anchor.column, to.row, to.column);
    const text = referenceText(range);
    const p = isPointing() ? state.current.point! : null;
    const start = p ? p.start : editor.selectionStart;
    const end = p ? p.start + p.length : editor.selectionEnd;
    editor.setRangeText(text, start, end, 'end');
    state.current.point = { anchor, cursor: to, range, start, length: text.length };
    host.draft?.(editor.value);
    refitEditor();
    reveal(to.row, to.column);
    refreshFormulaAids();
  };

  // --- navigation --------------------------------------------------------------------------------
  /** Widen a range until no merge straddles its edge (merges can chain, so to a fixed point). */
  const growOverMerges = (r: GridRange): GridRange => {
    const merges = host.sheet.value.merges;
    if (merges.length === 0) return r;
    let { top, left, bottom, right } = r;
    for (let changed = true; changed;) {
      changed = false;
      for (const m of merges) {
        const g = m.range;
        if (g.bottom < top || g.top > bottom || g.right < left || g.left > right) continue;
        if (g.top < top || g.left < left || g.bottom > bottom || g.right > right) {
          top = Math.min(top, g.top);
          left = Math.min(left, g.left);
          bottom = Math.max(bottom, g.bottom);
          right = Math.max(right, g.right);
          changed = true;
        }
      }
    }
    return { top, left, bottom, right };
  };

  const moveActive = (dr: number, dc: number, extend: boolean, cycleInRange = false, jump = false) => {
    const skipCell = (row: number, column: number) => {
      const v = state.current.viewport ?? viewport();
      if (v.rows.isHidden(row) || v.columns.isHidden(column)) return true;
      const m = host.sheet.value.findMerge(row, column);
      return !!m && (m.range.top !== row || m.range.left !== column);
    };
    if (cycleInRange && !extend && selection.cycle(dr, dc, skipCell)) {
      const a = selection.active.value;
      reveal(a.row, a.column);
      return;
    }
    const v = state.current.viewport ?? viewport();
    const sheet = host.sheet.value;
    const from = extend ? selection.extentCorner : selection.active.value;
    let row: number;
    let column: number;
    if (jump) ({ row, column } = host.edge(from.row, from.column, dr, dc));
    else if (extend) {
      row = dr ? v.rows.step(from.row, dr) : from.row;
      column = dc ? v.columns.step(from.column, dc) : from.column;
    } else if (Math.abs(dr) > 1) {
      // A page: count visible rows, so hidden ones don't shorten (or overshoot) it.
      row = v.rows.step(from.row, dr);
      column = from.column;
      if (v.rows.isHidden(row)) row = from.row;
    } else {
      [row, column] = sheet.stepOverMerges(from.row, from.column, dr, dc);
      // Past hidden lines; with nothing visible that way, stay put rather than land on a hidden one.
      if (dr && v.rows.isHidden(row)) row = v.rows.step(row, dr);
      if (dr && v.rows.isHidden(row)) row = from.row;
      if (dc && v.columns.isHidden(column)) column = v.columns.step(column, dc);
      if (dc && v.columns.isHidden(column)) column = from.column;
    }
    selection.select(row, column, { extend, grow: growOverMerges });
    reveal(row, column);
  };

  // The hover tip's pending timer must not fire into an unmounted grid.
  useEffect(() => () => clearTimeout(state.current.tipTimer), []);

  const pageRows = () => {
    const v = state.current.viewport ?? viewport();
    return Math.max(1, v.visibleRows.filter((b) => !b.frozen).length - 1);
  };

  const onKeyDown = (e: KeyboardEvent) => {
    if (editingRef.current) return;
    // An IME is composing (keyCode 229 where isComposing isn't set yet): compositionstart hands
    // it to the editor; the key itself is not a command or a typed character.
    if (e.isComposing || e.keyCode === 229) return;
    const mod = e.metaKey || e.ctrlKey;
    if (host.shortcut(e, false)) {
      e.preventDefault();
      return;
    }
    const arrows: Record<string, [number, number]> = {
      ArrowUp: [-1, 0],
      ArrowDown: [1, 0],
      ArrowLeft: [0, -1],
      ArrowRight: [0, 1],
    };
    if (e.key === ' ' && !mod && !e.altKey && host.toggleCheckboxes?.()) {
      e.preventDefault();
      return;
    }
    const arrow = arrows[e.key];
    if (arrow) {
      e.preventDefault();
      moveActive(arrow[0], arrow[1], e.shiftKey, false, mod);
      return;
    }
    switch (e.key) {
      case 'Enter':
        e.preventDefault();
        moveActive(e.shiftKey ? -1 : 1, 0, false, true);
        return;
      case 'Tab':
        e.preventDefault();
        moveActive(0, e.shiftKey ? -1 : 1, false, true);
        return;
      case 'PageDown':
      case 'PageUp':
        e.preventDefault();
        moveActive(e.key === 'PageDown' ? pageRows() : -pageRows(), 0, e.shiftKey);
        return;
      case 'Home': {
        e.preventDefault();
        // The first visible column (and row, with ⌘), not a hidden A or 1.
        const v = state.current.viewport ?? viewport();
        const firstVisible = (axis: typeof v.rows) => (axis.isHidden(0) ? axis.step(0, 1) : 0);
        const row = mod ? firstVisible(v.rows) : selection.active.value.row;
        const column = firstVisible(v.columns);
        selection.select(row, column, { extend: e.shiftKey });
        reveal(row, column);
        return;
      }
      case 'F2':
        e.preventDefault();
        beginEdit();
        return;
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        host.clearSelection();
        return;
      case 'Escape':
        host.clearMarquee?.();
        return;
    }
    // Typing a printable character replaces the cell's content and starts editing. ⌥ (macOS) and
    // AltGr (Ctrl+Alt elsewhere) type characters too; ⌘ and plain Ctrl are shortcuts.
    if (e.key.length === 1 && !e.metaKey && (!e.ctrlKey || e.altKey)) {
      e.preventDefault();
      beginEdit(e.key);
    }
  };

  // IME: composition starts in the grid → hand it to the editor so the candidate window works.
  const onCompositionStart = () => {
    if (!editingRef.current) beginEdit('');
  };

  const onEditorKeyDown = (e: KeyboardEvent) => {
    // The editor sits inside the grid's root: keep its keys from reaching the grid's handler.
    e.stopPropagation();
    if (e.isComposing || e.keyCode === 229) return;
    if (host.shortcut(e, true)) {
      e.preventDefault();
      return;
    }
    const plain = !e.altKey && !e.metaKey && !e.ctrlKey;
    // The function-name list takes ↑/↓ to choose, Tab to accept, and Escape to close.
    const list = assistRef.current?.suggestions ?? [];
    if (list.length > 0 && plain) {
      const at = assistRef.current!.index;
      if ((e.key === 'ArrowDown' || e.key === 'ArrowUp') && !e.shiftKey) {
        e.preventDefault();
        const step = e.key === 'ArrowDown' ? 1 : -1;
        setAssist({ ...assistRef.current!, index: (at + step + list.length) % list.length });
        return;
      }
      if (e.key === 'Tab' && !e.shiftKey) {
        e.preventDefault();
        acceptSuggestion(at);
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        setAssist(
          assistRef.current!.hint ? { suggestions: [], index: 0, hint: assistRef.current!.hint } : null,
        );
        return;
      }
    }
    // Pointing: in a formula where a reference can go, the arrows move a reference instead of
    // the caret (⇧ stretches it to a range).
    const arrows: Record<string, [number, number]> = {
      ArrowUp: [-1, 0],
      ArrowDown: [1, 0],
      ArrowLeft: [0, -1],
      ArrowRight: [0, 1],
    };
    const arrow = arrows[e.key];
    if (arrow && plain && canPoint()) {
      e.preventDefault();
      const at = editingRef.current!;
      const p = isPointing() ? state.current.point! : null;
      const from = p ? p.cursor : at;
      const v = state.current.viewport ?? viewport();
      const to = {
        row: arrow[0] ? v.rows.step(from.row, arrow[0]) : from.row,
        column: arrow[1] ? v.columns.step(from.column, arrow[1]) : from.column,
      };
      pointAt(e.shiftKey && p ? p.anchor : to, to);
      return;
    }
    if (e.key === 'Escape') {
      e.preventDefault();
      endEdit(false);
    } else if (e.key === 'Enter' && !e.shiftKey && !e.altKey) {
      e.preventDefault();
      endEdit(true, [1, 0]);
    } else if (e.key === 'Enter' && (e.altKey || e.shiftKey)) {
      // ⌥↩ / ⇧↩ start a new line in the cell. Inserted here rather than left to the textarea:
      // macOS doesn't reliably insert one for ⌥↩.
      e.preventDefault();
      insertLineBreak(e.currentTarget as HTMLTextAreaElement);
    } else if (e.key === 'Tab') {
      e.preventDefault();
      endEdit(true, [0, e.shiftKey ? -1 : 1]);
    } else if (
      (e.key === 'ArrowUp' || e.key === 'ArrowDown') &&
      !e.shiftKey &&
      !e.altKey &&
      !e.metaKey &&
      !e.ctrlKey
    ) {
      // ↑/↓ commit and move a cell, like Enter. In multi-line text they move between lines
      // first, and only leave the cell from the first or last line.
      const editor = e.currentTarget as HTMLTextAreaElement;
      const up = e.key === 'ArrowUp';
      const caret = editor.selectionStart;
      const onEdgeLine = up
        ? !editor.value.slice(0, caret).includes('\n')
        : !editor.value.slice(caret).includes('\n');
      if (!onEdgeLine) return;
      e.preventDefault();
      endEdit(true, [up ? -1 : 1, 0]);
    }
  };

  // --- pointer -----------------------------------------------------------------------------------
  const point = (e: PointerEvent) => {
    const r = canvasRef.current!.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  };

  const resizeTarget = (x: number, y: number): Drag | null => {
    const v = state.current.viewport;
    if (!v) return null;
    if (y < HEADER_HEIGHT && x > HEADER_WIDTH) {
      for (const b of v.visibleColumns)
        if (Math.abs(b.start + b.size - x) <= RESIZE_GRAB)
          return { kind: 'resizeColumn', index: b.index, origin: x, start: b.size / v.zoom };
    }
    if (x < HEADER_WIDTH && y > HEADER_HEIGHT) {
      for (const b of v.visibleRows)
        if (Math.abs(b.start + b.size - y) <= RESIZE_GRAB)
          return { kind: 'resizeRow', index: b.index, origin: y, start: b.size / v.zoom };
    }
    return null;
  };

  const handleAt = (x: number, y: number): boolean => {
    const v = state.current.viewport;
    if (!v || editingRef.current) return false;
    const h = fillHandleRect(v, selection.state.value);
    return (
      !!h &&
      x >= h.x - HANDLE_GRAB &&
      x <= h.x + h.w + HANDLE_GRAB &&
      y >= h.y - HANDLE_GRAB &&
      y <= h.y + h.h + HANDLE_GRAB
    );
  };

  /** The cursor for what is under the pointer: it says what a press there will do. */
  const cursorAt = (x: number, y: number, mod: boolean): string => {
    const v = state.current.viewport;
    if (!v) return 'default';
    const r = resizeTarget(x, y);
    if (r) return r.kind === 'resizeColumn' ? 'col-resize' : 'row-resize';
    if (hitChrome(state.current.chrome, x, y)) return 'pointer';
    if (handleAt(x, y)) return 'crosshair';
    const hit = v.hitTest(x, y);
    if (hit.zone !== 'cell') return 'default';
    if (mod && host.hasLink?.(hit.row, hit.column)) return 'pointer';
    if (checkboxAt(x, y)) return 'pointer';
    return 'cell';
  };

  /** The checkbox cell whose box is under (x, y), if any. A formula's box can't be clicked. */
  const checkboxAt = (x: number, y: number): { row: number; column: number } | null => {
    const v = state.current.viewport;
    if (!v || !host.toggleCheckboxes) return null;
    const hit = v.hitTest(x, y);
    if (hit.zone !== 'cell') return null;
    const sheet = host.sheet.value;
    const r = sheet.selectionRangeAt(hit.row, hit.column);
    const cell = sheet.find(r.top, r.left);
    if (!cell || cell.formula !== null || cell.value.kind !== 'boolean') return null;
    const style = host.styles.get(sheet.effectiveStyleId(r.top, r.left));
    if (!style.checkbox) return null;
    const box = checkboxRect(v.rangeRect(r.top, r.left, r.bottom, r.right), style, v.zoom);
    const slop = 3;
    const inside =
      x >= box.x - slop && x <= box.x + box.w + slop && y >= box.y - slop && y <= box.y + box.h + slop;
    return inside ? { row: r.top, column: r.left } : null;
  };

  const setCursor = (cursor: string) => {
    const el = scrollerRef.current;
    if (el && el.style.cursor !== cursor) el.style.cursor = cursor;
  };

  // ⌘/Ctrl going down or up over a link changes what a click does, so the cursor follows the key.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const p = state.current.pointer;
      if (!p || state.current.drag || (e.key !== 'Meta' && e.key !== 'Control')) return;
      setCursor(cursorAt(p.x, p.y, e.metaKey || e.ctrlKey));
    };
    addEventListener('keydown', onKey);
    addEventListener('keyup', onKey);
    return () => {
      removeEventListener('keydown', onKey);
      removeEventListener('keyup', onKey);
    };
  });

  // --- hover tooltip -------------------------------------------------------------------------------
  const hideTip = () => {
    const s = state.current;
    clearTimeout(s.tipTimer);
    s.tipCell = null;
    setTip(null);
  };

  const hoverTip = (x: number, y: number, clientX: number, clientY: number) => {
    const s = state.current;
    const v = s.viewport;
    if (!v || !host.hoverTip || editingRef.current) return hideTip();
    const hit = v.hitTest(x, y);
    if (hit.zone !== 'cell') return hideTip();
    const cell = host.sheet.value.selectionRangeAt(hit.row, hit.column);
    if (s.tipCell && s.tipCell.row === cell.top && s.tipCell.column === cell.left) return;
    hideTip();
    s.tipCell = { row: cell.top, column: cell.left };
    s.tipTimer = window.setTimeout(() => {
      const text = host.hoverTip?.(cell.top, cell.left);
      const r = rootRef.current!.getBoundingClientRect();
      if (text) setTip({ text, x: clientX - r.left + 12, y: clientY - r.top + 16 });
    }, TIP_DELAY_MS);
  };

  const onPointerDown = (e: PointerEvent) => {
    hideTip();
    if (e.button !== 0) return;
    // Ctrl-click on macOS is the context menu, not an added selection.
    if (IS_MAC && e.ctrlKey) return;
    const { x, y } = point(e);
    const v = state.current.viewport;
    if (!v) return;
    // Typing a formula where a reference can go: a click points at the cell instead of leaving.
    if (editingRef.current && editorRef.current!.value.startsWith('=')) {
      const hit = v.hitTest(x, y);
      if (hit.zone === 'cell' && (canPoint() || e.shiftKey)) {
        e.preventDefault();
        (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
        const p = isPointing() ? state.current.point : null;
        const to = { row: hit.row, column: hit.column };
        if (e.shiftKey && p) pointAt(p.anchor, to);
        else if (canPoint()) pointAt(to, to);
        state.current.drag = { kind: 'point', index: 0, origin: 0, start: 0 };
        editorRef.current!.focus();
        return;
      }
    }
    if (editingRef.current && !endEdit(true)) {
      // Rejected: stay in the editor on the same cell rather than moving the selection under it.
      e.preventDefault();
      editorRef.current?.focus();
      return;
    }
    rootRef.current!.focus();
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
    const resize = resizeTarget(x, y);
    if (resize) {
      if (e.detail === 2) {
        autofit(resize);
        return;
      }
      state.current.drag = resize;
      return;
    }
    if (handleAt(x, y)) {
      const source = selection.range.value;
      if (e.detail === 2) {
        const target = fillDownTarget(host.sheet.value, source);
        if (target) host.fill?.(source, target);
        return;
      }
      state.current.drag = { kind: 'fill', index: 0, origin: 0, start: 0, source };
      return;
    }
    const button = hitChrome(state.current.chrome, x, y);
    if (button && host.chrome) {
      const r = canvasRef.current!.getBoundingClientRect();
      host.chrome(button, {
        x: r.left + button.rect.x,
        y: r.top + button.rect.y,
        w: button.rect.w,
        h: button.rect.h,
      });
      return;
    }
    const hit = v.hitTest(x, y);
    const add = e.metaKey || e.ctrlKey;
    if (add && hit.zone === 'cell' && host.openLink?.(hit.row, hit.column)) return;
    switch (hit.zone) {
      case 'corner':
        selection.selectRange(rangeOf(0, 0, MAX_ROW, MAX_COLUMN));
        return;
      case 'columnHeader':
        if (e.shiftKey) selection.extendLines('columns', hit.column);
        else
          selection.selectRange(rangeOf(0, hit.column, MAX_ROW, hit.column), {
            row: v.visibleRows.find((b) => !b.frozen)?.index ?? 0,
            column: hit.column,
          });
        state.current.drag = { kind: 'columns', index: hit.column, origin: 0, start: 0 };
        return;
      case 'rowHeader':
        if (e.shiftKey) selection.extendLines('rows', hit.row);
        else selection.selectRange(rangeOf(hit.row, 0, hit.row, MAX_COLUMN));
        state.current.drag = { kind: 'rows', index: hit.row, origin: 0, start: 0 };
        return;
      case 'cell': {
        const box = !e.shiftKey && !add ? checkboxAt(x, y) : null;
        if (box) {
          // A click on the box ticks it (and selects the cell) without starting a drag.
          selection.select(box.row, box.column);
          host.toggleCheckboxes?.(box);
          return;
        }
        const merged = host.sheet.value.selectionRangeAt(hit.row, hit.column);
        if (e.shiftKey) selection.select(hit.row, hit.column, { extend: true, grow: growOverMerges });
        else selection.select(merged.top, merged.left, { add });
        state.current.drag = { kind: 'cells', index: 0, origin: 0, start: 0 };
        if (e.detail === 2) beginEdit();
        return;
      }
    }
  };

  /** Double-click on a header edge: fit the column to its content, or the row to its text. */
  const autofit = (edge: Drag) => {
    if (edge.kind === 'resizeColumn') host.autofitColumn?.(edge.index);
    // The row's own height goes back to the default; its content then sets how tall it draws.
    else if (edge.kind === 'resizeRow') host.resizeRow(edge.index, Defaults.rowHeight);
  };

  const onPointerMove = (e: PointerEvent) => {
    const { x, y } = point(e);
    const v = state.current.viewport;
    const drag = state.current.drag;
    state.current.pointer = { x, y };
    if (!drag) {
      setCursor(cursorAt(x, y, e.metaKey || e.ctrlKey));
      if (e.buttons === 0) hoverTip(x, y, e.clientX, e.clientY);
      return;
    }
    if (!v) return;
    switch (drag.kind) {
      case 'resizeColumn':
      case 'resizeRow': {
        const column = drag.kind === 'resizeColumn';
        const min = column ? MIN_COLUMN_WIDTH : MIN_ROW_HEIGHT;
        const size = Math.max(min, drag.start + ((column ? x : y) - drag.origin) / v.zoom);
        state.current.drag = { ...drag, size };
        const edge = drag.origin + (size - drag.start) * v.zoom;
        const g = guideRef.current;
        if (g)
          Object.assign(g.style, {
            display: 'block',
            left: column ? `${edge}px` : '0',
            top: column ? '0' : `${edge}px`,
            width: column ? '1px' : '100%',
            height: column ? '100%' : '1px',
          });
        return;
      }
      case 'cells': {
        const row = v.rowAt(y);
        const column = v.columnAt(x);
        selection.select(row, column, { extend: true, grow: growOverMerges });
        return;
      }
      case 'columns':
        selection.extendLines('columns', v.columnAt(x));
        return;
      case 'rows':
        selection.extendLines('rows', v.rowAt(y));
        return;
      case 'fill': {
        const target = fillTarget(drag.source!, v.rowAt(y), v.columnAt(x));
        const f = state.current.fill;
        if (
          !f ||
          f.top !== target.top ||
          f.left !== target.left ||
          f.bottom !== target.bottom ||
          f.right !== target.right
        ) {
          state.current.fill = target;
          draw();
        }
        return;
      }
      case 'point': {
        const p = state.current.point;
        if (!p) return;
        const to = { row: v.rowAt(y), column: v.columnAt(x) };
        if (to.row !== p.cursor.row || to.column !== p.cursor.column) pointAt(p.anchor, to);
        return;
      }
    }
  };

  const onPointerUp = () => {
    const drag = state.current.drag;
    state.current.drag = null;
    if (guideRef.current) guideRef.current.style.display = 'none';
    if (drag?.kind === 'fill') {
      const target = state.current.fill;
      state.current.fill = null;
      const src = drag.source!;
      if (
        target &&
        (target.top !== src.top ||
          target.left !== src.left ||
          target.bottom !== src.bottom ||
          target.right !== src.right)
      )
        host.fill?.(src, target);
      draw();
      return;
    }
    if (drag?.kind === 'point') {
      editorRef.current?.focus();
      return;
    }
    if (drag?.size === undefined) return;
    if (drag.kind === 'resizeColumn') host.resizeColumn(drag.index, drag.size);
    else if (drag.kind === 'resizeRow') host.resizeRow(drag.index, drag.size);
  };

  const onPointerLeave = () => {
    state.current.pointer = null;
    hideTip();
  };

  const onContextMenu = (e: MouseEvent) => {
    e.preventDefault();
    hideTip();
    if (!host.contextMenu) return;
    if (editingRef.current && !endEdit(true)) return;
    const r = canvasRef.current!.getBoundingClientRect();
    const v = state.current.viewport;
    if (!v) return;
    const hit = v.hitTest(e.clientX - r.left, e.clientY - r.top);
    if (hit.zone === 'corner') return;
    const ranges = selection.state.value.ranges;
    const inside = (row: number, column: number) =>
      ranges.some((g) => row >= g.top && row <= g.bottom && column >= g.left && column <= g.right);
    // A right-click outside the selection selects what was clicked first, so the menu acts on it.
    if (hit.zone === 'cell' && !inside(hit.row, hit.column)) {
      const m = host.sheet.value.selectionRangeAt(hit.row, hit.column);
      selection.select(m.top, m.left);
    } else if (
      hit.zone === 'columnHeader' &&
      !ranges.some(
        (g) => g.top === 0 && g.bottom === MAX_ROW && hit.column >= g.left && hit.column <= g.right,
      )
    )
      selection.selectRange(rangeOf(0, hit.column, MAX_ROW, hit.column), {
        row: v.visibleRows.find((b) => !b.frozen)?.index ?? 0,
        column: hit.column,
      });
    else if (
      hit.zone === 'rowHeader' &&
      !ranges.some((g) => g.left === 0 && g.right === MAX_COLUMN && hit.row >= g.top && hit.row <= g.bottom)
    )
      selection.selectRange(rangeOf(hit.row, 0, hit.row, MAX_COLUMN));
    rootRef.current?.focus();
    host.contextMenu(hit.zone, { x: e.clientX, y: e.clientY });
  };

  const onWheel = (e: WheelEvent) => {
    if (e.ctrlKey) return; // pinch-zoom is handled by the app's zoom command
    // Native scroller handles it; nothing to do.
  };

  const onEditorInput = (e: Event) => {
    const editor = e.currentTarget as HTMLTextAreaElement;
    const typed = (e as InputEvent).inputType === 'insertText';
    // Typing ends pointing: the next click or arrow inserts a new reference.
    state.current.point = null;
    // Type-ahead: appending to plain text offers the nearest entry above that starts with it,
    // selected, so the next keystroke replaces it and Enter keeps it.
    const at = editingRef.current;
    const text = editor.value;
    if (typed && at && !text.startsWith('=') && editor.selectionStart === text.length) {
      const match = columnCompletion(host.sheet.value, at.row, at.column, text);
      if (match) {
        editor.value = match;
        editor.setSelectionRange(text.length, match.length);
      }
    }
    host.draft?.(editor.value);
    refitEditor();
    refreshFormulaAids();
  };

  if (controller)
    controller.current = {
      focus: () => rootRef.current?.focus(),
      beginEdit,
      isEditing: () => editingRef.current !== null,
      commitEdit: () => endEdit(true),
    };

  return (
    <div
      ref={rootRef}
      class="cs-grid"
      tabIndex={0}
      role="grid"
      aria-label="Spreadsheet"
      onKeyDown={onKeyDown}
      onCompositionStart={onCompositionStart}
    >
      <canvas
        ref={canvasRef}
        class="cs-grid-canvas"
        style={{ width: `${size.w}px`, height: `${size.h}px` }}
      />
      <div
        ref={scrollerRef}
        class="cs-grid-scroller"
        onScroll={onScroll}
        onWheel={onWheel}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onPointerCancel={onPointerUp}
        onPointerLeave={onPointerLeave}
        onContextMenu={onContextMenu}
      >
        <div ref={spacerRef} class="cs-grid-spacer" />
      </div>
      <div ref={guideRef} class="cs-grid-resize-guide" />
      <textarea
        ref={editorRef}
        class="cs-grid-editor"
        spellcheck={false}
        autocomplete="off"
        onKeyDown={onEditorKeyDown}
        onKeyUp={(e) => {
          // The caret moved: the argument it is in may have changed.
          if (e.key.startsWith('Arrow') || e.key === 'Home' || e.key === 'End') refreshFormulaAids();
        }}
        onClick={refreshFormulaAids}
        onInput={onEditorInput}
        onBlur={(e) => {
          // Grid clicks are handled (and any rejection reported) by onPointerDown.
          if (rootRef.current?.contains(e.relatedTarget as Node | null)) return;
          endEdit(true, [0, 0], false);
        }}
      />
      {editing && assist && (
        <FormulaAssist assist={assist} editor={editorRef.current} onPick={acceptSuggestion} />
      )}
      {tip && (
        <div class="cs-grid-tip" role="tooltip" style={{ left: `${tip.x}px`, top: `${tip.y}px` }}>
          {tip.text}
        </div>
      )}
    </div>
  );
}

/** Function names matching what is typed, or the argument hint, just under the inline editor. */
function FormulaAssist(props: {
  assist: { suggestions: readonly FunctionMetadata[]; index: number; hint: readonly SignaturePart[] | null };
  editor: HTMLTextAreaElement | null;
  onPick(index: number): void;
}) {
  const { assist, editor } = props;
  if (!editor) return null;
  const left = parseFloat(editor.style.left) || 0;
  const top = (parseFloat(editor.style.top) || 0) + (parseFloat(editor.style.height) || 0) + 2;
  const style = { left: `${left}px`, top: `${top}px` };
  if (assist.suggestions.length > 0) {
    const current = assist.suggestions[assist.index];
    return (
      <div class="cs-grid-assist" style={style}>
        <div class="cs-grid-suggest" role="listbox" aria-label="Functions">
          {assist.suggestions.map((m, i) => (
            <div
              key={m.name}
              role="option"
              aria-selected={i === assist.index}
              class={i === assist.index ? 'on' : ''}
              onPointerDown={(e) => {
                // Keep the editor's focus; picking is a click, not a blur.
                e.preventDefault();
                props.onPick(i);
              }}
            >
              <span class="cs-grid-suggest-name">{m.name}</span>
            </div>
          ))}
        </div>
        {current && (
          <div class="cs-grid-suggest-about">
            <div class="cs-grid-suggest-sig">{current.signature}</div>
            <div>{current.summary}</div>
            <div class="cs-grid-suggest-key">Tab to insert</div>
          </div>
        )}
      </div>
    );
  }
  if (!assist.hint) return null;
  return (
    <div class="cs-grid-assist cs-grid-hint" style={style} role="status">
      {assist.hint.map((p, i) => (p.active ? <b key={i}>{p.text}</b> : <span key={i}>{p.text}</span>))}
    </div>
  );
}

/** Pointing while typing a formula: where the reference went in the text, and what it covers. */
interface Point {
  readonly anchor: { row: number; column: number };
  readonly cursor: { row: number; column: number };
  readonly range: GridRange;
  readonly start: number;
  readonly length: number;
}

interface Drag {
  readonly kind: 'resizeColumn' | 'resizeRow' | 'cells' | 'columns' | 'rows' | 'fill' | 'point';
  /** A fill drag's source range. */
  readonly source?: GridRange;
  readonly index: number;
  readonly origin: number;
  readonly start: number;
  /** The size a resize drag has reached so far. */
  readonly size?: number;
}
