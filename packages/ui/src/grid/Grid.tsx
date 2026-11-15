import { useEffect, useLayoutEffect, useRef, useState } from 'preact/hooks';
import { effect, type ReadonlySignal } from '@preact/signals';
import { Defaults, type CellStyle, type CellValue, type StylePool, type Worksheet } from '@cellestra/core';
import { chromeButtons, hitChrome, type ChromeButton } from './chrome.ts';
import { fillDownTarget, fillTarget } from './formulaEdit.ts';
import { InlineEdit, type CellAt, type FormulaAssist as Assist, type InlineEditDeps } from './inlineEdit.ts';
import {
  checkboxRect,
  fillHandleRect,
  renderGrid,
  MARQUEE_DASH,
  type DisplayText,
  type GridOverlay,
} from './render.ts';
import { CELL_PAD_X, CELL_PAD_Y } from './cellText.ts';
import { canvasRowMeasure, ContentHeights } from './rowHeights.ts';
import { afterTab, enterOrigin, startsComposition, type TabRun } from './keys.ts';
import { contentEnd, sheetLayout, type SheetLayout } from './sheetLayout.ts';
import { firstVisible, gridKeyAction, growOverMerges, stepFrom } from './navigation.ts';
import { insertLineBreak } from '../components/lineBreak.ts';
import { ClickCounter } from './clickCounter.ts';
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
   * The cells of `sheet` changed since `version` (edited, recalculated, or in a row or column an
   * edit restyled, resized or merged), or null when that is unknown. Lets the grid re-measure row
   * heights for just those cells.
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
  /** Scroll the active cell into view (after a jump the grid didn't make: Find, Go To). */
  revealActive(): void;
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

/** Whether the system asks for reduced motion (Settings › Accessibility › Display). */
const reducedMotion = () =>
  typeof matchMedia === 'function' && matchMedia('(prefers-reduced-motion: reduce)').matches;
const IS_MAC = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform);

export function Grid({ host, selection, zoom, palette, showFormulas, controller }: GridProps) {
  const rootRef = useRef<HTMLDivElement>(null);
  /** The line shown while dragging a column/row edge; the size is committed once, on release. */
  const guideRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const scrollerRef = useRef<HTMLDivElement>(null);
  const spacerRef = useRef<HTMLDivElement>(null);
  const editorRef = useRef<HTMLTextAreaElement>(null);
  const [size, setSize] = useState({ w: 0, h: 0 });
  const [editing, setEditing] = useState<CellAt | null>(null);
  const state = useRef({
    scrollX: 0,
    scrollY: 0,
    viewport: null as Viewport | null,
    /** The axes are rebuilt only when row/column layout moves (`layoutStamp`) or the sheet changes. */
    axesStamp: -1,
    axesSheet: null as Worksheet | null,
    axes: null as SheetLayout | null,
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
    /** Marching-ants dash offset. */
    phase: 0,
    /** The last pointer position over the grid (for the cursor when ⌘ goes down or up). */
    pointer: null as { x: number; y: number } | null,
    tipTimer: 0,
    tipCell: null as { row: number; column: number } | null,
    /** The Tab presses along the current row, so Enter can go back to where they started. */
    tabRun: null as TabRun | null,
  });
  const [tip, setTip] = useState<{ text: string; x: number; y: number } | null>(null);
  /** The formula aids under the editor: function names matching what is typed, or the argument hint. */
  const [assist, setAssist] = useState<Assist | null>(null);
  const clickCounter = useRef<ClickCounter | null>(null);
  // The functions the edit calls are declared below; it only calls them from handlers.
  const editDeps: InlineEditDeps = {
    host: {
      get sheet() {
        return host.sheet.value;
      },
      editText: (row, column) => host.editText(row, column),
      commit: (row, column, text) => host.commit(row, column, text),
      draft: (text) => host.draft?.(text),
    },
    get editor() {
      return editorRef.current;
    },
    view: {
      editingChanged: setEditing,
      assistChanged: setAssist,
      refit: () => refitEditor(),
      redraw: () => draw(),
      reveal: (row, column) => reveal(row, column),
    },
  };
  const editRef = useRef<InlineEdit | null>(null);
  const edit = (editRef.current ??= new InlineEdit(editDeps));
  edit.connect(editDeps);

  // --- layout -----------------------------------------------------------------------------------
  useLayoutEffect(() => {
    const el = rootRef.current!;
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const measureCtx = useRef<CanvasRenderingContext2D | null>(null);

  /** Recompute the rows content makes taller; true when that changed. */
  const refreshContentHeights = (sheet: Worksheet, version: number): boolean => {
    const s = state.current;
    const sameSheet = s.heights?.sheet === sheet;
    const sameLayout = s.contentStamp === sheet.layoutStamp;
    if (sameSheet && sameLayout && s.contentVersion === version) return false;
    const since = s.contentVersion;
    s.contentVersion = version;
    s.contentStamp = sheet.layoutStamp;
    // Same sheet: re-measure only the cells edited or recalculated since, and those in the rows and
    // columns an edit restyled, resized or merged. A layout change with no new version came from
    // outside the edits the session logs, so it measures everything again.
    const followed = sameSheet && (sameLayout || version !== since);
    const changed = followed ? host.cellsChangedSince?.(sheet.name, since) : null;
    if (changed && s.heights) return s.heights.update(changed);
    measureCtx.current ??= document.createElement('canvas').getContext('2d');
    s.heights = new ContentHeights(
      sheet,
      host.styles,
      canvasRowMeasure(sheet, (value, style) => host.display(value, style).text, measureCtx.current),
    );
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
      s.axes = sheetLayout(sheet, s.content);
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
      references: edit.references,
      point: edit.point?.range ?? null,
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

  // The marquee's ants march while it is up, unless the system asks for reduced motion: then the
  // dashes stand still.
  useEffect(() => {
    const m = host.marquee;
    if (!m || reducedMotion()) return;
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

  /** Scroll so the cell shows; `axes` limits it to one direction (a whole-row selection's far end). */
  const reveal = (row: number, column: number, axes = { x: true, y: true }) => {
    const v = state.current.viewport ?? viewport();
    const x = axes.x ? v.scrollToReveal('x', column) : state.current.scrollX;
    const y = axes.y ? v.scrollToReveal('y', row) : state.current.scrollY;
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
    edit.begin({ row, column }, initial, () => {
      host.clearMarquee?.();
      hideTip();
      // Show and focus synchronously, inside the key event: the next keystrokes (and an IME
      // composition) must land in the editor, not the grid.
      reveal(row, column);
      placeEditor(state.current.viewport ?? viewport());
      editorRef.current!.focus();
    });
  };

  /** False when the host rejected the value (a stop validation rule); the editor stays open. */
  const endEdit = (commit: boolean, move: readonly [number, number] = [0, 0], refocus = true): boolean => {
    const result = edit.end(commit);
    if (result === 'refused') {
      if (refocus) editorRef.current!.select();
      return false;
    }
    if (result === 'none') return true;
    if (refocus) rootRef.current?.focus();
    if (move[0] || move[1]) moveActive(move[0], move[1], false, true);
    return true;
  };

  const placeEditor = (v: Viewport) => {
    const editor = editorRef.current;
    if (!editor) return;
    const editing = edit.at;
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
    if (v && edit.at) placeEditor(v);
  };

  // --- navigation --------------------------------------------------------------------------------
  const growOverSheetMerges = (r: GridRange): GridRange => growOverMerges(host.sheet.value.merges, r);

  const moveActive = (dr: number, dc: number, extend: boolean, cycleInRange = false, jump = false) => {
    const skipCell = (row: number, column: number) => {
      const v = state.current.viewport ?? viewport();
      if (v.rows.isHidden(row) || v.columns.isHidden(column)) return true;
      const m = host.sheet.value.findMerge(row, column);
      return !!m && (m.range.top !== row || m.range.left !== column);
    };
    const run = state.current.tabRun;
    state.current.tabRun = null;
    if (cycleInRange && !extend && selection.cycle(dr, dc, skipCell)) {
      const a = selection.active.value;
      reveal(a.row, a.column);
      return;
    }
    const v = state.current.viewport ?? viewport();
    const sheet = host.sheet.value;
    const active = selection.active.value;
    // Enter after a run of Tabs goes down from the column the run started in.
    const from = extend
      ? selection.extentCorner
      : cycleInRange && dr === 1 && dc === 0
        ? enterOrigin(run, active)
        : active;
    const { row, column } = jump
      ? host.edge(from.row, from.column, dr, dc)
      : stepFrom(from, dr, dc, extend ? 'extend' : Math.abs(dr) > 1 ? 'page' : 'cell', v, sheet);
    selection.select(row, column, { extend, grow: growOverSheetMerges });
    if (cycleInRange && !extend && dr === 0 && dc !== 0)
      state.current.tabRun = afterTab(run, active, selection.active.value);
    // Extending a whole-row (or whole-column) selection must not scroll to its far end.
    reveal(row, column, extend ? { x: dc !== 0, y: dr !== 0 } : undefined);
  };

  // The hover tip's pending timer must not fire into an unmounted grid.
  useEffect(() => () => clearTimeout(state.current.tipTimer), []);

  const pageRows = () => {
    const v = state.current.viewport ?? viewport();
    return Math.max(1, v.visibleRows.filter((b) => !b.frozen).length - 1);
  };

  const onKeyDown = (e: KeyboardEvent) => {
    if (edit.at) return;
    // A dead key or an IME: its character arrives later as input, and a grid that can't take text
    // would lose it. Open the editor now, inside this key event and without preventing it, so the
    // composition lands there. The key itself is not a command.
    if (startsComposition(e)) {
      beginEdit('');
      edit.composing = true;
      return;
    }
    const mod = e.metaKey || e.ctrlKey;
    if (host.shortcut(e, false)) {
      e.preventDefault();
      return;
    }
    if (e.key === ' ' && !mod && !e.altKey && host.toggleCheckboxes?.()) {
      e.preventDefault();
      return;
    }
    const action = gridKeyAction(e);
    if (!action) return;
    if (action.kind === 'escape') {
      // Esc drops the copy marquee first; with none, it collapses the selection to the active cell.
      if (host.marquee?.value?.sheet === host.sheet.value.name) {
        host.clearMarquee?.();
        return;
      }
      const { row, column } = selection.active.value;
      const m = host.sheet.value.selectionRangeAt(row, column);
      if (selection.collapse({ top: m.top, left: m.left, bottom: m.bottom, right: m.right }))
        e.preventDefault();
      return;
    }
    e.preventDefault();
    switch (action.kind) {
      case 'move':
        moveActive(action.rows, action.columns, action.extend, action.cycle, action.jump);
        return;
      case 'page':
        moveActive(action.down ? pageRows() : -pageRows(), 0, action.extend);
        return;
      case 'home': {
        // The first visible column (and row, with ⌘), not a hidden A or 1.
        const v = state.current.viewport ?? viewport();
        const row = action.top ? firstVisible(v.rows) : selection.active.value.row;
        const column = firstVisible(v.columns);
        selection.select(row, column, { extend: action.extend });
        reveal(row, column);
        return;
      }
      case 'edit':
        beginEdit();
        return;
      case 'clear':
        host.clearSelection();
        return;
      case 'type':
        beginEdit(action.text);
        return;
    }
  };

  // IME: composition starts in the grid → hand it to the editor so the candidate window works.
  const onCompositionStart = () => {
    if (edit.at) return;
    beginEdit('');
    edit.composing = true;
  };

  const onEditorKeyDown = (e: KeyboardEvent) => {
    // The editor sits inside the grid's root: keep its keys from reaching the grid's handler.
    e.stopPropagation();
    if (e.isComposing || e.keyCode === 229) return;
    if (host.shortcut(e, true)) {
      e.preventDefault();
      return;
    }
    const action = edit.key(e, () => state.current.viewport ?? viewport());
    if (!action) return;
    e.preventDefault();
    if (action.kind === 'end') endEdit(action.commit, action.move);
    // Inserted here rather than left to the textarea: macOS doesn't reliably insert one for ⌥↩.
    else if (action.kind === 'newline') insertLineBreak(e.currentTarget as HTMLTextAreaElement);
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
    if (!v || edit.at) return false;
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
    if (!v || !host.hoverTip || edit.at) return hideTip();
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
    const clicks = (clickCounter.current ??= new ClickCounter()).press(e.timeStamp, x, y, e.detail);
    const v = state.current.viewport;
    if (!v) return;
    // Typing a formula where a reference can go: a click points at the cell instead of leaving.
    if (edit.at && editorRef.current!.value.startsWith('=')) {
      const hit = v.hitTest(x, y);
      if (hit.zone === 'cell' && (edit.canPoint() || e.shiftKey)) {
        e.preventDefault();
        (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
        const p = edit.activePoint;
        const to = { row: hit.row, column: hit.column };
        if (e.shiftKey && p) edit.pointAt(p.anchor, to);
        else if (edit.canPoint()) edit.pointAt(to, to);
        state.current.drag = { kind: 'point', index: 0, origin: 0, start: 0 };
        editorRef.current!.focus();
        return;
      }
    }
    if (edit.at && !endEdit(true)) {
      // Rejected: stay in the editor on the same cell rather than moving the selection under it.
      e.preventDefault();
      editorRef.current?.focus();
      return;
    }
    rootRef.current!.focus();
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
    const resize = resizeTarget(x, y);
    if (resize) {
      if (clicks === 2) {
        autofit(resize);
        return;
      }
      state.current.drag = resize;
      return;
    }
    if (handleAt(x, y)) {
      const source = selection.range.value;
      if (clicks === 2) {
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
        if (e.shiftKey) selection.select(hit.row, hit.column, { extend: true, grow: growOverSheetMerges });
        else selection.select(merged.top, merged.left, { add });
        state.current.drag = { kind: 'cells', index: 0, origin: 0, start: 0 };
        if (clicks === 2) beginEdit();
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
        selection.select(row, column, { extend: true, grow: growOverSheetMerges });
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
        const p = edit.point;
        if (!p) return;
        const to = { row: v.rowAt(y), column: v.columnAt(x) };
        if (to.row !== p.cursor.row || to.column !== p.cursor.column) edit.pointAt(p.anchor, to);
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
    if (edit.at && !endEdit(true)) return;
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

  const onEditorInput = (e: Event) => edit.input((e as InputEvent).inputType === 'insertText');

  if (controller)
    controller.current = {
      focus: () => rootRef.current?.focus(),
      beginEdit,
      isEditing: () => edit.at !== null,
      revealActive: () => {
        const { row, column } = selection.active.peek();
        reveal(row, column);
      },
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
          if (e.key.startsWith('Arrow') || e.key === 'Home' || e.key === 'End') edit.refreshAids();
        }}
        onClick={() => edit.refreshAids()}
        onInput={onEditorInput}
        onBlur={(e) => {
          // Grid clicks are handled (and any rejection reported) by onPointerDown.
          if (rootRef.current?.contains(e.relatedTarget as Node | null)) return;
          // The window lost focus (⌘Tab to look something up): the edit stays open, and the browser
          // gives the editor its focus back when the window returns.
          if (!e.currentTarget.ownerDocument.hasFocus()) return;
          endEdit(true, [0, 0], false);
        }}
      />
      {editing && assist && (
        <FormulaAssist assist={assist} editor={editorRef.current} onPick={(i) => edit.acceptSuggestion(i)} />
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
  assist: Assist;
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
