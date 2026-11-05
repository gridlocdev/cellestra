import { useEffect, useLayoutEffect, useRef, useState } from 'preact/hooks';
import { effect, type ReadonlySignal } from '@preact/signals';
import type { CellStyle, CellValue, StylePool, Worksheet } from '@cellestra/core';
import { Axis } from './axis.ts';
import { renderGrid, CELL_PAD_X, CELL_PAD_Y, type DisplayText } from './render.ts';
import { Viewport, HEADER_HEIGHT, HEADER_WIDTH } from './viewport.ts';
import type { GridPalette } from './palette.ts';
import { MAX_COLUMN, MAX_ROW, rangeOf, type Selection } from '../session/selection.ts';

/** What the grid needs from the document. The grid never edits the model itself. */
export interface GridHost {
  readonly sheet: ReadonlySignal<Worksheet>;
  readonly styles: StylePool;
  /** Bumped on every model change; the grid redraws when it moves. */
  readonly version: ReadonlySignal<number>;
  display(value: CellValue, style: CellStyle): DisplayText;
  /** What the inline editor starts with (`=SUM(A1:A3)` for formulas, the raw input otherwise). */
  editText(row: number, column: number): string;
  commit(row: number, column: number, text: string): void;
  clearSelection(): void;
  /** Edge navigation (⌘-arrow): the cell to jump to. */
  edge(row: number, column: number, rowStep: number, columnStep: number): { row: number; column: number };
  /** Key not handled by the grid (shortcuts) → command catalog. Return true when handled. */
  shortcut(e: KeyboardEvent, inText: boolean): boolean;
  /** Column/row resize from the header borders. */
  resizeColumn(column: number, width: number): void;
  resizeRow(row: number, height: number): void;
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

function axesFor(sheet: Worksheet): { columns: Axis; rows: Axis } {
  const cols: [number, number][] = [];
  for (const [i, c] of sheet.columns) cols.push([i, c.hidden ? 0 : c.width]);
  const rows: [number, number][] = [];
  for (const [i, r] of sheet.rows) rows.push([i, r.hidden ? 0 : r.height]);
  return {
    columns: new Axis(MAX_COLUMN + 1, 64, cols),
    rows: new Axis(MAX_ROW + 1, 20, rows),
  };
}

export function Grid({ host, selection, zoom, palette, showFormulas, controller }: GridProps) {
  const rootRef = useRef<HTMLDivElement>(null);
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
    axesVersion: -1,
    axes: null as ReturnType<typeof axesFor> | null,
    drag: null as null | Drag,
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
    if (!s.axes || s.axesVersion !== version) {
      s.axes = axesFor(sheet);
      s.axesVersion = version;
    }
    return new Viewport({
      ...s.axes,
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
    renderGrid({
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
    const used = host.sheet.value.getUsedRange();
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
  const syncingScroll = useRef(false);
  const onScroll = () => {
    if (syncingScroll.current) {
      syncingScroll.current = false;
      return;
    }
    const el = scrollerRef.current!;
    state.current.scrollX = el.scrollLeft / zoom.value;
    state.current.scrollY = el.scrollTop / zoom.value;
    draw();
  };

  const reveal = (row: number, column: number) => {
    const v = state.current.viewport ?? viewport();
    const x = v.scrollToReveal('x', column);
    const y = v.scrollToReveal('y', row);
    if (x !== state.current.scrollX || y !== state.current.scrollY) {
      state.current.scrollX = x;
      state.current.scrollY = y;
      const el = scrollerRef.current!;
      draw();
      syncingScroll.current = true;
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

  const endEdit = (commit: boolean, move: [number, number] = [0, 0]) => {
    const at = editingRef.current;
    if (!at) return;
    const text = editorRef.current!.value;
    setEditing(null);
    if (commit) host.commit(at.row, at.column, text);
    rootRef.current?.focus();
    if (move[0] || move[1]) moveActive(move[0], move[1], false, true);
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
  const moveActive = (dr: number, dc: number, extend: boolean, cycleInRange = false, jump = false) => {
    if (cycleInRange && !extend && selection.cycle(dr, dc)) {
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
      if (dr && v.rows.isHidden(row)) row = v.rows.step(row, dr);
      if (dc && v.columns.isHidden(column)) column = v.columns.step(column, dc);
    }
    selection.select(row, column, { extend });
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
    if (editingRef.current) endEdit(true);
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
    const hit = v.hitTest(x, y);
    const add = e.metaKey || e.ctrlKey;
    switch (hit.zone) {
      case 'corner':
        selection.selectRange(rangeOf(0, 0, MAX_ROW, MAX_COLUMN));
        return;
      case 'columnHeader':
        if (e.shiftKey) selection.select(MAX_ROW, hit.column, { extend: true });
        else
          selection.selectRange(rangeOf(0, hit.column, MAX_ROW, hit.column), {
            row: v.visibleRows.find((b) => !b.frozen)?.index ?? 0,
            column: hit.column,
          });
        state.current.drag = { kind: 'columns', index: hit.column, origin: 0, start: 0 };
        return;
      case 'rowHeader':
        if (e.shiftKey) selection.select(hit.row, MAX_COLUMN, { extend: true });
        else selection.selectRange(rangeOf(hit.row, 0, hit.row, MAX_COLUMN));
        state.current.drag = { kind: 'rows', index: hit.row, origin: 0, start: 0 };
        return;
      case 'cell': {
        const merged = host.sheet.value.selectionRangeAt(hit.row, hit.column);
        selection.select(merged.top, merged.left, { extend: e.shiftKey, add });
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
        host.resizeColumn(drag.index, Math.max(MIN_COLUMN_WIDTH, drag.start + (x - drag.origin) / v.zoom));
        return;
      case 'resizeRow':
        host.resizeRow(drag.index, Math.max(MIN_ROW_HEIGHT, drag.start + (y - drag.origin) / v.zoom));
        return;
      case 'cells': {
        const row = v.rowAt(y);
        const column = v.columnAt(x);
        selection.select(row, column, { extend: true });
        return;
      }
      case 'columns':
        selection.select(MAX_ROW, v.columnAt(x), { extend: true });
        return;
      case 'rows':
        selection.select(v.rowAt(y), MAX_COLUMN, { extend: true });
        return;
    }
  };

  const onPointerUp = () => {
    state.current.drag = null;
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
      <textarea
        ref={editorRef}
        class="cs-grid-editor"
        spellcheck={false}
        autocomplete="off"
        onKeyDown={onEditorKeyDown}
        onBlur={() => endEdit(true)}
      />
    </div>
  );
}

interface Drag {
  readonly kind: 'resizeColumn' | 'resizeRow' | 'cells' | 'columns' | 'rows';
  readonly index: number;
  readonly origin: number;
  readonly start: number;
}
