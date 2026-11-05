import { useEffect, useLayoutEffect, useRef, useState } from 'preact/hooks';
import { effect, type ReadonlySignal } from '@preact/signals';
import type { CellStyle, CellValue, StylePool, Worksheet } from '@cellestra/core';
import { Axis } from './axis.ts';
import { chromeButtons, hitChrome, type ChromeButton } from './chrome.ts';
import { renderGrid, CELL_PAD_X, CELL_PAD_Y, type DisplayText } from './render.ts';
import { Viewport, HEADER_HEIGHT, HEADER_WIDTH } from './viewport.ts';
import type { GridPalette } from './palette.ts';
import { MAX_COLUMN, MAX_ROW, rangeOf, type GridRange, type Selection } from '../session/selection.ts';

/** What the grid needs from the document. The grid never edits the model itself. */
export interface GridHost {
  readonly sheet: ReadonlySignal<Worksheet>;
  readonly styles: StylePool;
  /** Bumped on every model change; the grid redraws when it moves. */
  readonly version: ReadonlySignal<number>;
  display(value: CellValue, style: CellStyle): DisplayText;
  /** What the inline editor starts with (`=SUM(A1:A3)` for formulas, the raw input otherwise). */
  editText(row: number, column: number): string;
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
}

const MIN_COLUMN_WIDTH = 16;
const MIN_ROW_HEIGHT = 8;
const RESIZE_GRAB = 4;

interface Axes {
  readonly columns: Axis;
  readonly rows: Axis;
  /** The last row and column with non-default layout (−1 for none), for the scroll extent. */
  readonly lastRow: number;
  readonly lastColumn: number;
}

function axesFor(sheet: Worksheet): Axes {
  const cols: [number, number][] = [];
  let lastColumn = -1;
  for (const [i, c] of sheet.columns) {
    cols.push([i, c.hidden ? 0 : c.width]);
    if (!c.isDefault) lastColumn = Math.max(lastColumn, i);
  }
  const rows: [number, number][] = [];
  let lastRow = -1;
  for (const [i, r] of sheet.rows) {
    rows.push([i, r.hidden ? 0 : r.height]);
    if (!r.isDefault) lastRow = Math.max(lastRow, i);
  }
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
  });

  // --- layout -----------------------------------------------------------------------------------
  useLayoutEffect(() => {
    const el = rootRef.current!;
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const viewport = (): Viewport => {
    const s = state.current;
    const version = host.version.value;
    const sheet = host.sheet.value;
    if (!s.axes || s.axesStamp !== sheet.layoutStamp || s.axesSheet !== sheet) {
      s.axes = axesFor(sheet);
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
    renderGrid({
      chrome,
      ctx,
      viewport: v,
      sheet: host.sheet.value,
      styles: host.styles,
      palette: palette.value,
      selection: sel,
      display: (val, style) => host.display(val, style),
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
    // Show and focus synchronously, inside the key event: the next keystrokes (and an IME
    // composition) must land in the editor, not the grid.
    reveal(row, column);
    placeEditor(state.current.viewport ?? viewport());
    editor.focus();
    const end = editor.value.length;
    editor.setSelectionRange(end, end);
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
    const style = host.styles.get(host.sheet.value.find(editing.row, editing.column)?.styleId ?? 0);
    Object.assign(editor.style, {
      display: 'block',
      left: `${rect.x}px`,
      top: `${rect.y}px`,
      minWidth: `${rect.w}px`,
      minHeight: `${rect.h}px`,
      padding: `${CELL_PAD_Y * v.zoom}px ${CELL_PAD_X * v.zoom}px`,
      font: `${style.font.italic ? 'italic ' : ''}${style.font.bold ? 'bold ' : ''}${((style.font.sizePoints * 96) / 72) * v.zoom}px "${style.font.family}", "Helvetica Neue", Arial, sans-serif`,
    });
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

  const pageRows = () => {
    const v = state.current.viewport ?? viewport();
    return Math.max(1, v.visibleRows.filter((b) => !b.frozen).length - 1);
  };

  const onKeyDown = (e: KeyboardEvent) => {
    if (editingRef.current) return;
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
      case 'Home':
        e.preventDefault();
        selection.select(mod ? 0 : selection.active.value.row, 0, { extend: e.shiftKey });
        reveal(mod ? 0 : selection.active.value.row, 0);
        return;
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
        return;
    }
    // Typing a printable character replaces the cell's content and starts editing.
    if (e.key.length === 1 && !mod && !e.altKey) {
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
    if (e.isComposing) return;
    if (host.shortcut(e, true)) {
      e.preventDefault();
      return;
    }
    if (e.key === 'Escape') {
      e.preventDefault();
      endEdit(false);
    } else if (e.key === 'Enter' && !e.shiftKey && !e.altKey) {
      e.preventDefault();
      endEdit(true, [1, 0]);
    } else if (e.key === 'Enter' && (e.altKey || e.shiftKey)) {
      // ⌥↩ / ⇧↩ insert a line break — the textarea's default.
    } else if (e.key === 'Tab') {
      e.preventDefault();
      endEdit(true, [0, e.shiftKey ? -1 : 1]);
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

  const onPointerDown = (e: PointerEvent) => {
    if (e.button !== 0) return;
    if (editingRef.current && !endEdit(true)) {
      // Rejected: stay in the editor on the same cell rather than moving the selection under it.
      e.preventDefault();
      editorRef.current?.focus();
      return;
    }
    const { x, y } = point(e);
    const v = state.current.viewport;
    if (!v) return;
    rootRef.current!.focus();
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
    const resize = resizeTarget(x, y);
    if (resize) {
      state.current.drag = resize;
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
        const merged = host.sheet.value.selectionRangeAt(hit.row, hit.column);
        if (e.shiftKey) selection.select(hit.row, hit.column, { extend: true, grow: growOverMerges });
        else selection.select(merged.top, merged.left, { add });
        state.current.drag = { kind: 'cells', index: 0, origin: 0, start: 0 };
        if (e.detail === 2) beginEdit();
        return;
      }
    }
  };

  const onPointerMove = (e: PointerEvent) => {
    const { x, y } = point(e);
    const v = state.current.viewport;
    const drag = state.current.drag;
    const canvas = canvasRef.current!;
    if (!drag) {
      const r = resizeTarget(x, y);
      canvas.style.cursor = r ? (r.kind === 'resizeColumn' ? 'col-resize' : 'row-resize') : 'cell';
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
    }
  };

  const onPointerUp = () => {
    const drag = state.current.drag;
    state.current.drag = null;
    if (guideRef.current) guideRef.current.style.display = 'none';
    if (drag?.size === undefined) return;
    if (drag.kind === 'resizeColumn') host.resizeColumn(drag.index, drag.size);
    else if (drag.kind === 'resizeRow') host.resizeRow(drag.index, drag.size);
  };

  const onWheel = (e: WheelEvent) => {
    if (e.ctrlKey) return; // pinch-zoom is handled by the app's zoom command
    // Native scroller handles it; nothing to do.
  };

  if (controller)
    controller.current = {
      focus: () => rootRef.current?.focus(),
      beginEdit,
      isEditing: () => editingRef.current !== null,
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
        onBlur={(e) => {
          // Grid clicks are handled (and any rejection reported) by onPointerDown.
          if (rootRef.current?.contains(e.relatedTarget as Node | null)) return;
          endEdit(true, [0, 0], false);
        }}
      />
    </div>
  );
}

interface Drag {
  readonly kind: 'resizeColumn' | 'resizeRow' | 'cells' | 'columns' | 'rows';
  readonly index: number;
  readonly origin: number;
  readonly start: number;
  /** The size a resize drag has reached so far. */
  readonly size?: number;
}
