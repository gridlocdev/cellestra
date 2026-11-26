import { useComputed, useSignal, useSignalEffect, type ReadonlySignal } from '@preact/signals';
import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'preact/hooks';
import { SheetDimensions, type CellValue } from '@cellestra/core';
import {
  CommandMenu,
  Dialog,
  FormatCellsDialog,
  ConditionalFormatDialog,
  PageSetupDialog,
  FormulaBar,
  Grid,
  NasBrowser,
  QuickPick,
  SheetTabs,
  after,
  calculationBadge,
  saveStateText,
  type CommandCatalog,
  type GridHost,
  type RecentEntry,
} from '@cellestra/ui';
import {
  AutoSaveSeconds,
  autoSaveFromInput,
  wheelZoomFactor,
  type AppController,
  type Layout,
  type Theme,
} from './controller.ts';
import {
  FilterPopup,
  FunctionPicker,
  HyperlinkDialog,
  NamedRangesDialog,
  SortFilterDialog,
  ValidationDialog,
  ValidationListPopup,
} from './DataDialogs.tsx';
import { ChartPane } from './ChartPane.tsx';
import { cellAnnouncement } from './announce.ts';
import { Ribbon } from './Ribbon.tsx';
import { GridContextMenu } from './ContextMenu.tsx';
import { ShortcutsDialog } from './ShortcutsDialog.tsx';
import { Toolbar } from './Toolbar.tsx';

export function App({ app, catalog }: { readonly app: AppController; readonly catalog: CommandCatalog }) {
  const s = app.session;
  const dragging = useSignal(false);
  /** The cell editor's text while editing: the formula bar shows it as it is typed. */
  const draft = useSignal<string | null>(null);

  const host = useMemo<GridHost>(
    () => ({
      sheet: s.activeSheet,
      get styles() {
        return s.workbook.value.styles;
      },
      version: s.version,
      cellsChangedSince: (sheet, version) => s.cellsChangedSince(sheet, version),
      display: (value: CellValue, style) => ({ text: s.display(value, style) }),
      conditionalStyle: (style, row, column) => app.conditionalStyle(style, row, column),
      editText: (row, column) => s.editText(row, column),
      sheetsBetween: (first, last) =>
        s.workbook
          .peek()
          .sheetsBetween(first, last)
          ?.map((sheet) => sheet.name) ?? null,
      draft: (text) => (draft.value = text),
      commit: (row, column, text) => app.commit(row, column, text),
      clearSelection: () => void app.clearContents(),
      edge: (row, column, dr, dc) => app.edge(row, column, dr, dc),
      shortcut: (e, inText) => {
        // Picking cells for a chart: Enter takes the selection, Escape goes back unchanged.
        if (app.chartPick.peek() && !inText && (e.key === 'Enter' || e.key === 'Escape')) {
          if (e.key === 'Enter') app.finishChartPick();
          else app.cancelChartPick();
          return true;
        }
        // A dialog or menu open over the grid keeps the shortcuts to itself.
        if (app.modalOpen()) return false;
        // Copy/cut/paste use the native clipboard events below (no permission prompt).
        const command = catalog.forEvent(e, inText);
        if (!command || ['edit.copy', 'edit.cut', 'edit.paste'].includes(command.id)) return false;
        void catalog.run(command.id);
        return true;
      },
      resizeColumn: (column, width) => app.resizeColumn(column, width),
      resizeRow: (row, height) => app.resizeRow(row, height),
      chrome: (button, anchor) => {
        if (button.kind === 'filter') {
          const target = app.sheet.tables.find((t) => t.range.equals(button.range)) ?? null;
          app.dialog.value = {
            kind: 'filter',
            target: { range: button.range, hasHeader: button.hasHeader, table: target },
            column: button.column,
            anchor,
          };
        } else {
          const { row, column } = button;
          app.selection.select(row, column);
          app.dialog.value = { kind: 'validationList', rule: button.rule, row, column, anchor };
        }
      },
      openLink: (row, column) => app.openLink(row, column),
      hasLink: (row, column) => !!app.sheet.find(row, column)?.hyperlink,
      toggleCheckboxes: (at) => app.toggleCheckboxes(at),
      fill: (source, target) => app.fillRange(source, target),
      marquee: app.marquee,
      clearMarquee: () => app.clearMarquee(),
      stylePreview: app.stylePreview,
      autofitColumn: (column) => app.autofitColumn(column),
      contextMenu: (zone, at) => (app.contextMenu.value = { zone, ...at }),
      hoverTip: (row, column) => app.hoverTip(row, column),
    }),
    [app, catalog, s, draft],
  );

  // Pinching a trackpad or ⌘/Ctrl-scrolling zooms the sheet, not the whole page (ribbon included).
  useEffect(() => {
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      app.zoomBy(wheelZoomFactor(e.deltaY, e.deltaMode));
    };
    // Safari reports a pinch as gesture events with a running scale, not as ⌘-wheel.
    let lastScale = 1;
    const onGestureStart = (e: Event) => {
      e.preventDefault();
      lastScale = 1;
    };
    const onGestureChange = (e: Event) => {
      e.preventDefault();
      const scale = (e as Event & { scale?: number }).scale ?? 1;
      if (scale > 0 && lastScale > 0) app.zoomBy(scale / lastScale);
      lastScale = scale;
    };
    const options = { passive: false } as const;
    addEventListener('wheel', onWheel, options);
    addEventListener('gesturestart', onGestureStart, options);
    addEventListener('gesturechange', onGestureChange, options);
    return () => {
      removeEventListener('wheel', onWheel);
      removeEventListener('gesturestart', onGestureStart);
      removeEventListener('gesturechange', onGestureChange);
    };
  }, [app]);

  // Shortcuts that fire anywhere (outside the grid too): ⌘K, ⌘S, ⌘O…
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.defaultPrevented || app.modalOpen()) return;
      const target = e.target as HTMLElement | null;
      const inText = !!target?.closest('input, textarea, select, [contenteditable]');
      const inGrid = !!target?.closest('.cs-grid');
      if (inGrid) return; // the grid routes its own keys through host.shortcut
      const command = catalog.forEvent(e, inText);
      if (!command || (inText && !command.global)) return;
      if (['edit.copy', 'edit.cut', 'edit.paste'].includes(command.id)) return;
      e.preventDefault();
      void catalog.run(command.id);
    };
    const inField = (e: Event) =>
      e.target instanceof Element && !!e.target.closest('input, textarea:not(.cs-grid-editor)');
    const gridHasFocus = () =>
      !!document.activeElement?.closest('.cs-grid') && !(app.grid.current?.isEditing() ?? false);
    const onCopy = (e: ClipboardEvent) => {
      if (inField(e) || !gridHasFocus()) return;
      e.preventDefault();
      app.copy(e.type === 'cut', e.clipboardData);
    };
    const onPaste = (e: ClipboardEvent) => {
      if (inField(e) || !gridHasFocus()) return;
      e.preventDefault();
      app.paste(e.clipboardData?.getData('text/plain') ?? null);
    };
    // The browser's own ⌘Z undoes the last typing in any field, even one without focus (from a
    // dialog's button, say): it would put old text back into a hidden editor as a draft.
    const onHistory = (e: InputEvent) => {
      if (e.inputType.startsWith('history') && e.target !== document.activeElement) e.preventDefault();
    };
    addEventListener('keydown', onKey);
    document.addEventListener('beforeinput', onHistory, true);
    document.addEventListener('copy', onCopy);
    document.addEventListener('cut', onCopy);
    document.addEventListener('paste', onPaste);
    return () => {
      removeEventListener('keydown', onKey);
      document.removeEventListener('beforeinput', onHistory, true);
      document.removeEventListener('copy', onCopy);
      document.removeEventListener('cut', onCopy);
      document.removeEventListener('paste', onPaste);
    };
  }, [app, catalog]);

  // Window title.
  useSignalEffect(() => {
    document.title = `${s.modified.value ? '• ' : ''}${s.title.value} — Cellestra`;
  });

  // Drag a file anywhere to open it.
  useEffect(() => {
    let depth = 0;
    const hasFiles = (e: DragEvent) => e.dataTransfer?.types.includes('Files') ?? false;
    const enter = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      depth++;
      dragging.value = true;
    };
    const leave = () => {
      depth = Math.max(0, depth - 1);
      if (depth === 0) dragging.value = false;
    };
    const over = (e: DragEvent) => {
      if (hasFiles(e)) e.preventDefault();
    };
    const drop = (e: DragEvent) => {
      depth = 0;
      dragging.value = false;
      const file = e.dataTransfer?.files[0];
      if (!file) return;
      e.preventDefault();
      // Chromium can hand over a writable handle, so Save goes back to the dropped file. It must
      // be asked for now: the DataTransfer is emptied once this handler returns.
      const item = e.dataTransfer.items[0] as
        (DataTransferItem & { getAsFileSystemHandle?: () => Promise<FileSystemHandle | null> }) | undefined;
      void app.openFile(file, item?.getAsFileSystemHandle?.());
    };
    addEventListener('dragenter', enter);
    addEventListener('dragleave', leave);
    addEventListener('dragover', over);
    addEventListener('drop', drop);
    return () => {
      removeEventListener('dragenter', enter);
      removeEventListener('dragleave', leave);
      removeEventListener('dragover', over);
      removeEventListener('drop', drop);
    };
  }, [app, dragging]);

  const isMacNative = app.native?.platform === 'darwin';
  const ribbon = app.layout.value === 'ribbon';

  // The ribbon's top bar is the (taller) title bar, so the window controls move down into it.
  useEffect(() => {
    if (isMacNative) app.native?.setTitleBarHeight?.(ribbon ? 44 : 28);
  }, [app, isMacNative, ribbon]);

  return (
    <div class={`cs-app ${isMacNative ? 'cs-native-mac' : ''}`}>
      {ribbon ? (
        <Ribbon app={app} catalog={catalog} macTitleBar={isMacNative} />
      ) : (
        <>
          {isMacNative && <TitleBar app={app} />}
          <Toolbar app={app} catalog={catalog} titleInline={!isMacNative} />
        </>
      )}
      <FormulaBarConnected app={app} draft={draft} />
      {app.chartPick.value && <ChartPickBanner app={app} />}
      <main class="cs-main" onPointerUp={() => app.selectionDone()}>
        {s.activeSheet.value.isChartSheet ? (
          <ChartPane app={app} />
        ) : (
          <Grid
            host={host}
            selection={app.selection}
            zoom={app.gridZoom}
            palette={app.palette}
            showFormulas={app.showFormulas}
            controller={app.grid}
          />
        )}
      </main>
      <GridContextMenu app={app} catalog={catalog} />
      <Footer app={app} />
      <Overlays app={app} catalog={catalog} />
      {dragging.value && <div class="cs-drop">Drop a spreadsheet to open it</div>}
      <Announcer app={app} />
      {app.toast.value && (
        <div
          class="cs-toast"
          style={
            app.toast.value.tone === 'error' ? { background: 'var(--cs-danger)', color: '#fff' } : undefined
          }
        >
          {/* The announcer reads the text; the Undo button is for keyboard and screen-reader users. */}
          <span aria-hidden="true">{app.toast.value.text}</span>
          {app.noticeCanUndo && (
            <button type="button" class="cs-toast-undo" onClick={() => app.undoFromNotice()}>
              Undo
            </button>
          )}
        </div>
      )}
    </div>
  );
}

/**
 * The Mac window's title bar: the workbook's name centred on the row with the window controls,
 * where the eye finds it first. Elsewhere the native title bar (or the browser tab) shows it.
 */
function TitleBar({ app }: { readonly app: AppController }) {
  const s = app.session;
  return (
    <div class="cs-titlebar">
      <span class="cs-title">
        {s.modified.value ? '• ' : ''}
        {s.title.value}
      </span>
    </div>
  );
}

const cellLabel = (row: number, column: number) => SheetDimensions.columnToLetters(column) + (row + 1);

/** `A1:C9` for a range of more than one cell, null for a single cell. */
function rangeLabel(r: { top: number; left: number; bottom: number; right: number }): string | null {
  if (r.top === r.bottom && r.left === r.right) return null;
  return `${cellLabel(r.top, r.left)}:${cellLabel(r.bottom, r.right)}`;
}

/**
 * Screen-reader live regions, visually hidden: one says the active cell's address and value each
 * time it moves (or its value changes), one repeats toasts. They stay in the page so the reader
 * notices their text changing; a region that appears with its text already in it is often missed.
 */
function Announcer({ app }: { readonly app: AppController }) {
  const s = app.session;
  const cell = useComputed(() => {
    void s.version.value;
    if (s.activeSheet.value.isChartSheet) return '';
    const { row, column } = app.selection.active.value;
    const range = rangeLabel(app.selection.range.value);
    return cellAnnouncement(cellLabel(row, column), app.shownText(row, column), range);
  });
  const toast = app.toast.value;
  return (
    <>
      <div class="cs-sr-only" aria-live="polite" aria-atomic="true">
        {cell.value}
      </div>
      <div class="cs-sr-only" role="status" aria-atomic="true">
        {toast?.tone === 'info' ? toast.text : ''}
      </div>
      <div class="cs-sr-only" role="alert" aria-atomic="true">
        {toast?.tone === 'error' ? toast.text : ''}
      </div>
    </>
  );
}

function FormulaBarConnected({
  app,
  draft,
}: {
  readonly app: AppController;
  readonly draft: ReadonlySignal<string | null>;
}) {
  const s = app.session;
  const address = useComputed(() => {
    const st = app.selection.state.value;
    return rangeLabel(app.selection.range.value) ?? cellLabel(st.row, st.column);
  });
  // The cell typing in the bar belongs to: its document, sheet and position.
  const target = useComputed(() => {
    void s.workbook.value;
    const { row, column } = app.selection.active.value;
    return `${s.generation}|${s.activeSheet.value.name}|${row}|${column}`;
  });
  const content = useComputed(() => {
    void s.version.value;
    void s.activeSheet.value;
    const typing = draft.value;
    if (typing !== null) return typing;
    const { row, column } = app.selection.active.value;
    return s.editText(row, column);
  });
  return (
    <FormulaBar
      disabled={s.activeSheet.value.isChartSheet}
      address={address.value}
      content={content.value}
      target={target.value}
      onCommit={(text) => {
        const { row, column } = app.selection.active.peek();
        if (!app.commit(row, column, text)) return;
        app.selection.select(Math.min(row + 1, 1_048_575), column);
      }}
      nameBoxRef={app.nameBox}
      onGoTo={(ref) => app.goTo(ref)}
      onDone={() => app.grid.current?.focus()}
    />
  );
}

function Footer({ app }: { readonly app: AppController }) {
  const s = app.session;
  const [menu, setMenu] = useState<{ index: number; x: number; y: number } | null>(null);
  const summary = useSignal('');
  useSignalEffect(() => {
    void s.version.value;
    const scope = summaryScope(app);
    if (!scope) {
      summary.value = '';
      return;
    }
    if (Math.min(scope.area, scope.sheet.cellCount) <= SUMMARY_SYNC_CELLS) {
      summary.value = selectionSummary(app, scope);
      return;
    }
    return after(120, () => (summary.value = selectionSummary(app, scope)));
  });
  const inputMessage = useComputed(() => {
    void s.version.value;
    const { row, column } = app.selection.active.value;
    return s.activeSheet.value.findValidation(row, column)?.inputMessage ?? '';
  });
  const saveState = useComputed(() =>
    saveStateText(s.modified.value, s.document.value !== null, s.savedAt.value),
  );
  return (
    <div class="cs-footer">
      <SheetTabs
        names={s.sheetNames.value}
        active={s.activeIndex.value}
        onSelect={(i, via) => {
          // An open cell edit belongs to the sheet it was typed on.
          if (app.grid.current?.commitEdit() ?? true) s.setActiveSheet(i);
          // An arrow key along the tabs keeps the keyboard there.
          if (via === 'keyboard') return;
          // The tab button took focus with the click: hand the keyboard back to the grid so arrows
          // and typing act on the new sheet's active cell.
          queueMicrotask(() => app.grid.current?.focus());
        }}
        onAdd={() => {
          if (!(app.grid.current?.commitEdit() ?? true)) return;
          s.addSheet();
          queueMicrotask(() => app.grid.current?.focus());
        }}
        onRename={(i, name) => s.renameSheet(i, name)}
        onMove={(from, to) => s.moveSheet(from, to)}
        onMenu={(index, x, y) => setMenu({ index, x, y })}
      />
      {s.lastAction.value && (
        <span class="cs-status cs-last-action" title={s.lastAction.value}>
          {s.lastAction.value}
        </span>
      )}
      {inputMessage.value && (
        <span class="cs-status cs-input-message" title={inputMessage.value}>
          {inputMessage.value}
        </span>
      )}
      {s.hasCircularReference.value && (
        <span class="cs-status" style={{ color: 'var(--cs-danger)' }}>
          Circular reference
        </span>
      )}
      <span class="cs-status">{summary.value}</span>
      <button
        type="button"
        class="cs-status cs-calc-badge"
        title={
          s.calculationMode.value === 'manual'
            ? 'Manual calculation: press F9 to recalculate. Click to switch to automatic.'
            : 'Automatic calculation. Click to switch to manual.'
        }
        onClick={() => {
          s.setCalculationMode(s.calculationMode.peek() === 'manual' ? 'automatic' : 'manual');
          queueMicrotask(() => app.grid.current?.focus());
        }}
      >
        {calculationBadge(s.calculationMode.value)}
      </button>
      <span class="cs-status" title="Zoom">
        {Math.round(app.zoom.value * 100)}%
      </span>
      {saveState.value && <span class="cs-status">{saveState.value}</span>}
      {menu && (
        <SheetMenu
          x={menu.x}
          y={menu.y}
          canDelete={s.sheetNames.value.length > 1}
          onClose={() => setMenu(null)}
          onPick={(action) => {
            setMenu(null);
            if (action === 'rename') void app.renameSheet(menu.index);
            else if (action === 'duplicate') s.duplicateSheet(menu.index);
            else void app.deleteSheet(menu.index);
          }}
        />
      )}
    </div>
  );
}

function SheetMenu(props: {
  x: number;
  y: number;
  canDelete: boolean;
  onClose: () => void;
  onPick: (a: 'rename' | 'duplicate' | 'delete') => void;
}) {
  useEffect(() => {
    const close = () => props.onClose();
    const cancel = after(0, () => addEventListener('pointerdown', close, { once: true }));
    return () => {
      cancel();
      removeEventListener('pointerdown', close);
    };
  }, []);
  // Open above the tab that was clicked (the tab bar is at the bottom), kept inside the window.
  const ref = useRef<HTMLDivElement>(null);
  const [at, setAt] = useState<{ x: number; y: number } | null>(null);
  useLayoutEffect(() => {
    const el = ref.current;
    const w = el?.offsetWidth ?? 180;
    const h = el?.offsetHeight ?? 100;
    setAt({
      x: Math.max(4, Math.min(props.x, innerWidth - w - 4)),
      y: Math.max(4, props.y - h - 4),
    });
  }, [props.x, props.y]);
  const item = (a: 'rename' | 'duplicate' | 'delete', label: string, disabled = false) => (
    <button
      type="button"
      role="menuitem"
      class="cs-context-item"
      disabled={disabled}
      onPointerDown={(e) => {
        // Keep the window listener from closing the menu before the click lands.
        e.stopPropagation();
        e.preventDefault();
      }}
      onClick={() => props.onPick(a)}
    >
      <span class="cs-context-title">{label}</span>
    </button>
  );
  return (
    <div
      ref={ref}
      class="cs-context-menu"
      role="menu"
      aria-label="Sheet actions"
      style={{
        left: `${(at ?? props).x}px`,
        top: `${(at ?? props).y}px`,
        minWidth: '180px',
        visibility: at ? 'visible' : 'hidden',
      }}
      onContextMenu={(e) => e.preventDefault()}
    >
      {item('rename', 'Rename…')}
      {item('duplicate', 'Duplicate')}
      {item('delete', 'Delete', !props.canDelete)}
    </div>
  );
}

/** The selection's ranges cut to the sheet's last stored cell, and how many cells they cover. */
function summaryScope(app: AppController) {
  const st = app.selection.state.value;
  const sheet = app.session.activeSheet.value;
  const r = app.selection.range.value;
  const extent = sheet.cellExtent();
  if ((r.top === r.bottom && r.left === r.right && st.ranges.length === 1) || !extent) return null;
  const ranges = st.ranges.map((range) => ({
    top: range.top,
    left: range.left,
    bottom: Math.min(range.bottom, extent.lastRow),
    right: Math.min(range.right, extent.lastColumn),
  }));
  const area = ranges.reduce(
    (n, g) => n + Math.max(0, g.bottom - g.top + 1) * Math.max(0, g.right - g.left + 1),
    0,
  );
  return { sheet, ranges, area };
}

/** The prompt while cells are picked for a chart: what to select, and how to finish. */
function ChartPickBanner({ app }: { readonly app: AppController }) {
  const pick = app.chartPick.value;
  if (!pick) return null;
  return (
    <div class="cs-pick-banner" role="status">
      <span>
        {pick.field === 'data'
          ? `Select the data for the chart “${pick.chartSheet.name}”: labels in the first column, a series in each other column. Then press Enter. Esc cancels.`
          : `Select the cells for ${pick.label} of the chart “${pick.chartSheet.name}”, then press Enter. Esc cancels.`}
      </span>
      <button type="button" class="cs-btn" onClick={() => app.finishChartPick()}>
        Use selection
      </button>
      <button type="button" class="cs-btn" onClick={() => app.cancelChartPick()}>
        Cancel
      </button>
    </div>
  );
}

/** Sum / average / count of the selection's numbers, like the old status bar. */
function selectionSummary(app: AppController, scope: NonNullable<ReturnType<typeof summaryScope>>): string {
  const { sheet, ranges, area } = scope;
  let count = 0;
  let numbers = 0;
  let sum = 0;
  const add = (v: CellValue) => {
    if (v.kind === 'empty') return;
    count++;
    if (v.kind === 'number') {
      numbers++;
      sum += v.value;
    }
  };
  if (area > sheet.cellCount) {
    // A big selection over sparse data: visit the stored cells, not every position.
    for (const [row, column, cell] of sheet.enumerateCells()) {
      // Like Calc: rows a filter (or the user) hid don't count.
      if (sheet.findRow(row)?.hidden) continue;
      for (const g of ranges)
        if (row >= g.top && row <= g.bottom && column >= g.left && column <= g.right) add(cell.value);
    }
  } else {
    for (const g of ranges) {
      for (let row = g.top; row <= g.bottom; row++) {
        if (sheet.findRow(row)?.hidden) continue;
        for (let c = g.left; c <= g.right; c++) add(sheet.getValue(row, c));
      }
    }
  }
  if (count === 0) return '';
  const fmt = (n: number) => n.toLocaleString(app.session.locale, { maximumFractionDigits: 10 });
  return numbers > 0 ? `Sum ${fmt(sum)}   Average ${fmt(sum / numbers)}   Count ${count}` : `Count ${count}`;
}

/** Above this many cells the summary waits for the selection to settle instead of following each move. */
const SUMMARY_SYNC_CELLS = 20_000;

/** Focus the grid when nothing in particular has it, as after a dialog whose opener has gone. */
function focusGridIfLost(app: AppController): void {
  const active = document.activeElement;
  if (!active || active === document.body || !active.isConnected) app.grid.current?.focus();
}

function Overlays({ app, catalog }: { readonly app: AppController; readonly catalog: CommandCatalog }) {
  const s = app.session;
  const d = app.dialog.value;
  // A dialog hands focus back to whatever opened it (a sheet tab, a ribbon tab, the grid). When that
  // was a menu item that has since gone (or nothing), focus would be left on the page itself, where
  // no key reaches the grid: use the grid.
  const refocus = () => queueMicrotask(() => focusGridIfLost(app));
  const open = d !== null;
  const wasOpen = useRef(open);
  useEffect(() => {
    const closed = wasOpen.current && !open;
    wasOpen.current = open;
    if (closed) focusGridIfLost(app);
  }, [open]);
  return (
    <>
      {app.paletteOpen.value && (
        <CommandMenu
          catalog={catalog}
          onClose={() => {
            app.paletteOpen.value = false;
            refocus();
          }}
        />
      )}
      {app.sheetPickerOpen.value && (
        <QuickPick
          items={s.sheetNames.value.map((name, index) => ({ name, index }))}
          label={(i) => i.name}
          name="Go to sheet"
          placeholder="Go to sheet…"
          onPick={(i) => s.setActiveSheet(i.index)}
          onClose={() => {
            app.sheetPickerOpen.value = false;
            refocus();
          }}
        />
      )}
      {app.recentOpen.value && (
        <QuickPick<RecentEntry>
          items={app.recents.value}
          label={(e) => `${e.ref.name} — ${storeLabel(e.ref.store)}`}
          name="Open recent"
          placeholder="Open recent…"
          onPick={(e) => void app.openRecent(e)}
          onClose={() => {
            app.recentOpen.value = false;
            refocus();
          }}
        />
      )}
      {d?.kind === 'confirm-discard' && (
        <Dialog
          title="Unsaved changes"
          onClose={() => d.resolve('cancel')}
          actions={
            <>
              <button type="button" class="cs-btn danger" onClick={() => d.resolve('discard')}>
                Don’t Save
              </button>
              <button type="button" class="cs-btn" onClick={() => d.resolve('cancel')}>
                Cancel
              </button>
              <button type="submit" class="cs-btn primary" autofocus>
                Save
              </button>
            </>
          }
          onSubmit={() => d.resolve('save')}
        >
          <p>Do you want to save the changes to “{s.title.value}”?</p>
        </Dialog>
      )}
      {d?.kind === 'conflict' && (
        <Dialog
          title={
            d.reason === 'exists'
              ? 'File exists'
              : d.reason === 'locked'
                ? 'File is open elsewhere'
                : 'File changed'
          }
          onClose={() => d.resolve('cancel')}
          actions={
            <>
              {d.reason !== 'exists' && (
                <button type="button" class="cs-btn" onClick={() => d.resolve('reload')}>
                  Discard Mine &amp; Reload
                </button>
              )}
              {d.reason !== 'exists' && (
                <button type="button" class="cs-btn" onClick={() => d.resolve('saveAs')}>
                  Save a Copy…
                </button>
              )}
              <button type="button" class="cs-btn" onClick={() => d.resolve('cancel')}>
                Cancel
              </button>
              <button type="button" class="cs-btn danger" onClick={() => d.resolve('overwrite')}>
                {d.reason === 'exists' ? 'Replace' : 'Overwrite'}
              </button>
            </>
          }
        >
          <p>{d.message}</p>
          {d.reason === 'changed' && (
            <p>Someone else saved this file after you opened it. Overwriting replaces their changes.</p>
          )}
          {d.reason === 'locked' && (
            <p>LibreOffice has the file open. Overwriting may lose the changes made there.</p>
          )}
        </Dialog>
      )}
      {d?.kind === 'outside-change' && (
        <Dialog
          title="File changed on the server"
          onClose={() => d.resolve(false)}
          actions={
            <>
              <button type="button" class="cs-btn" onClick={() => d.resolve(false)}>
                Keep This Version
              </button>
              <button type="submit" class="cs-btn primary" autofocus>
                Reload
              </button>
            </>
          }
          onSubmit={() => d.resolve(true)}
        >
          <p>“{d.name}” was saved on the server since you opened it. Reload it to see those changes?</p>
          <p>If you keep this version, saving it later asks before overwriting theirs.</p>
        </Dialog>
      )}
      {d?.kind === 'nas' && (
        <NasBrowser
          store={app.nas}
          mode={d.mode}
          {...(d.suggestedName === undefined ? {} : { suggestedName: d.suggestedName })}
          onDone={(r) => d.resolve(r)}
        />
      )}
      {d?.kind === 'prompt' && (
        <PromptDialog title={d.title} label={d.label} value={d.value} onDone={d.resolve} />
      )}
      {d?.kind === 'confirm' && (
        <Dialog
          title={d.title}
          onClose={() => d.resolve(false)}
          actions={
            <>
              <button type="button" class="cs-btn" onClick={() => d.resolve(false)}>
                Cancel
              </button>
              <button type="submit" class="cs-btn primary" autofocus>
                {d.ok}
              </button>
            </>
          }
          onSubmit={() => d.resolve(true)}
        >
          <p>{d.message}</p>
        </Dialog>
      )}
      {d?.kind === 'recovery' && (
        <Dialog
          title="Recover unsaved work?"
          onClose={() => d.resolve(false)}
          actions={
            <>
              <button type="button" class="cs-btn" onClick={() => d.resolve(false)}>
                Discard
              </button>
              <button type="submit" class="cs-btn primary" autofocus>
                Recover
              </button>
            </>
          }
          onSubmit={() => d.resolve(true)}
        >
          <p>Cellestra closed with unsaved changes to “{d.name}”. Recover them?</p>
        </Dialog>
      )}
      {d?.kind === 'csv-loss' && (
        <Dialog
          title="Keep CSV format?"
          onClose={() => d.resolve('cancel')}
          actions={
            <>
              <button type="button" class="cs-btn" onClick={() => d.resolve('cancel')}>
                Cancel
              </button>
              <button type="button" class="cs-btn" onClick={() => d.resolve('csv')}>
                Keep CSV Format
              </button>
              <button type="submit" class="cs-btn primary" autofocus>
                Save as .ods…
              </button>
            </>
          }
          onSubmit={() => d.resolve('ods')}
        >
          <p>
            CSV keeps only the values of the current sheet. Saving “{d.name}” as CSV loses its{' '}
            {d.losses.join(', ')}.
          </p>
        </Dialog>
      )}
      {d?.kind === 'find' && <FindDialog app={app} />}
      {d?.kind === 'shortcuts' && (
        <ShortcutsDialog
          catalog={catalog}
          native={app.native}
          onClose={() => {
            app.dialog.value = null;
            refocus();
          }}
        />
      )}
      {d?.kind === 'about' && (
        <Dialog
          title="About Cellestra"
          onClose={() => {
            app.dialog.value = null;
            refocus();
          }}
          actions={
            <button type="submit" class="cs-btn primary">
              OK
            </button>
          }
          onSubmit={() => {
            app.dialog.value = null;
            refocus();
          }}
        >
          <p>
            <strong>Cellestra</strong> — a spreadsheet for OpenDocument (.ods) files.
          </p>
          <p>
            {app.native
              ? 'A desktop app that works on files on this computer.'
              : app.nasAvailable.value
                ? 'Runs in your browser, with files in folders on your NAS.'
                : 'Runs entirely in your browser; files stay on this computer.'}{' '}
            MIT licensed.
          </p>
        </Dialog>
      )}
      {d?.kind === 'formatCells' && (
        <FormatCellsDialog
          current={app.activeStyle}
          styles={d.styles}
          locale={s.locale}
          onApply={(transform) => app.style('Format Cells', transform)}
          onClose={() => {
            app.dialog.value = null;
            refocus();
          }}
        />
      )}
      {d?.kind === 'pageSetup' && (
        <PageSetupDialog
          sheetName={app.sheetName}
          current={app.sheet.pageSetup}
          onApply={(setup) => app.setPageSetup(setup)}
          onPrint={(setup) => void app.print(false, setup)}
          onClose={() => {
            app.dialog.value = null;
            refocus();
          }}
        />
      )}
      {d?.kind === 'conditionalFormats' && (
        <ConditionalFormatDialog
          current={app.activeStyle}
          base={{ sheet: app.sheetName, row: app.ranges[0]!.top, column: app.ranges[0]!.left }}
          onApply={(transform) => app.style('Conditional Formatting', transform)}
          onClose={() => {
            app.dialog.value = null;
            refocus();
          }}
        />
      )}
      {d?.kind === 'filter' && (
        <FilterPopup
          key={`${d.target.range.toString()}:${d.column}`}
          app={app}
          target={d.target}
          column={d.column}
          anchor={d.anchor}
        />
      )}
      {d?.kind === 'validationList' && (
        <ValidationListPopup
          key={`${d.row}:${d.column}`}
          app={app}
          rule={d.rule}
          row={d.row}
          column={d.column}
          anchor={d.anchor}
        />
      )}
      {d?.kind === 'sortFilter' && <SortFilterDialog app={app} target={d.target} sortOnly={d.sortOnly} />}
      {d?.kind === 'validation' && <ValidationDialog app={app} range={d.range} existing={d.existing} />}
      {d?.kind === 'hyperlink' && <HyperlinkDialog app={app} row={d.row} column={d.column} />}
      {d?.kind === 'namedRanges' && <NamedRangesDialog app={app} />}
      {d?.kind === 'functions' && <FunctionPicker app={app} />}
      {d?.kind === 'settings' && (
        <SettingsDialog
          app={app}
          onClose={() => {
            app.dialog.value = null;
            refocus();
          }}
        />
      )}
    </>
  );
}

function storeLabel(kind: string): string {
  return kind === 'nas' ? 'NAS' : kind === 'native' ? 'This Mac' : 'This computer';
}

function PromptDialog(props: {
  title: string;
  label: string;
  value: string;
  onDone: (v: string | null) => void;
}) {
  const [value, setValue] = useState(props.value);
  return (
    <Dialog
      title={props.title}
      onClose={() => props.onDone(null)}
      onSubmit={() => props.onDone(value)}
      actions={
        <>
          <button type="button" class="cs-btn" onClick={() => props.onDone(null)}>
            Cancel
          </button>
          <button type="submit" class="cs-btn primary">
            OK
          </button>
        </>
      }
    >
      <label style={{ display: 'grid', gap: '6px' }}>
        {props.label}
        <input
          value={value}
          autofocus
          onInput={(e) => setValue(e.currentTarget.value)}
          onFocus={(e) => e.currentTarget.select()}
        />
      </label>
    </Dialog>
  );
}

/** How many matches the Find dialog lists; the count still says how many there are. */
const FIND_LIST_LIMIT = 200;

function FindDialog({ app }: { readonly app: AppController }) {
  const [query, setQuery] = useState(app.findQuery.peek());
  const [replacement, setReplacement] = useState('');
  const [note, setNote] = useState('');
  // The list follows edits (Replace changes it) and the sheet in view.
  const version = app.session.version.value;
  const sheet = app.session.activeSheet.value;
  const active = app.selection.active.value;
  const matches = useMemo(() => app.findAll(query), [app, query, version, sheet]);
  const close = () => {
    app.findQuery.value = query;
    app.dialog.value = null;
    queueMicrotask(() => focusGridIfLost(app));
  };
  const next = (back = false) => {
    app.findQuery.value = query;
    setNote(app.findNext(query, back) ? '' : 'No matches.');
  };
  const replaceOne = () => {
    app.findQuery.value = query;
    const r = app.replaceNext(query, replacement);
    setNote(r.found ? '' : r.replaced ? 'Replaced the last match.' : 'No matches.');
  };
  const count = matches.length === 1 ? '1 match' : `${matches.length} matches`;
  return (
    <Dialog
      title="Find and Replace"
      onClose={close}
      onSubmit={() => next()}
      actions={
        <>
          <button
            type="button"
            class="cs-btn"
            onClick={() => setNote(`Replaced ${app.replaceAll(query, replacement)} cell(s).`)}
          >
            Replace All
          </button>
          <button type="button" class="cs-btn" onClick={replaceOne}>
            Replace
          </button>
          <button type="button" class="cs-btn" onClick={() => next(true)}>
            Previous
          </button>
          <button type="submit" class="cs-btn primary">
            Find Next
          </button>
        </>
      }
    >
      <div style={{ display: 'grid', gap: '8px', minWidth: '320px' }}>
        <label style={{ display: 'grid', gap: '4px' }}>
          Find
          <input value={query} autofocus onInput={(e) => setQuery(e.currentTarget.value)} />
        </label>
        <label style={{ display: 'grid', gap: '4px' }}>
          Replace with
          <input value={replacement} onInput={(e) => setReplacement(e.currentTarget.value)} />
        </label>
        <span class="cs-status" style={{ padding: 0 }} aria-live="polite">
          {note || (query ? count : '')}
        </span>
        {matches.length > 0 && (
          <ul class="cs-find-results" aria-label="Matches">
            {matches.slice(0, FIND_LIST_LIMIT).map((m) => (
              <li key={`${m.row}:${m.column}`}>
                <button
                  type="button"
                  aria-current={m.row === active.row && m.column === active.column ? 'true' : undefined}
                  onClick={() => {
                    app.findQuery.value = query;
                    app.goToMatch(m);
                  }}
                >
                  <span class="cs-find-address">{m.address}</span>
                  <span class="cs-find-text">{m.text}</span>
                </button>
              </li>
            ))}
            {matches.length > FIND_LIST_LIMIT && (
              <li class="cs-muted">and {matches.length - FIND_LIST_LIMIT} more</li>
            )}
          </ul>
        )}
      </div>
    </Dialog>
  );
}

function SettingsDialog({ app, onClose }: { readonly app: AppController; readonly onClose: () => void }) {
  const themes: Theme[] = ['system', 'light', 'dark'];
  return (
    <Dialog
      title="Settings"
      onClose={onClose}
      actions={
        <button type="submit" class="cs-btn primary">
          Done
        </button>
      }
      onSubmit={onClose}
    >
      <div style={{ display: 'grid', gap: '12px', minWidth: '320px' }}>
        <label style={{ display: 'grid', gap: '4px' }}>
          Appearance
          <select
            value={app.theme.value}
            onChange={(e) => (app.theme.value = e.currentTarget.value as Theme)}
          >
            {themes.map((t) => (
              <option key={t} value={t}>
                {t === 'system' ? 'Match system' : t[0]!.toUpperCase() + t.slice(1)}
              </option>
            ))}
          </select>
        </label>
        <label style={{ display: 'grid', gap: '4px' }}>
          Layout
          <select
            value={app.layout.value}
            onChange={(e) => (app.layout.value = e.currentTarget.value as Layout)}
          >
            <option value="ribbon">Ribbon (tabs)</option>
            <option value="bar">Compact toolbar</option>
          </select>
        </label>
        <label style={{ display: 'grid', gap: '4px' }}>
          Auto-save (seconds)
          <input
            type="number"
            min={0}
            max={AutoSaveSeconds.maximum}
            step={5}
            value={app.autoSaveSeconds.value}
            onChange={(e) => {
              const v = autoSaveFromInput(e.currentTarget.value, app.autoSaveSeconds.peek());
              app.autoSaveSeconds.value = v;
              e.currentTarget.value = String(v);
            }}
          />
          <span class="cs-muted" style={{ fontSize: '12px' }}>
            Seconds of no editing before the file is saved in place. 0 turns auto-save off; anything below{' '}
            {AutoSaveSeconds.minimum} counts as {AutoSaveSeconds.minimum}. New and uploaded files wait for ⌘S.
          </span>
        </label>
        {app.nasAvailable.value && (
          <label style={{ display: 'grid', gap: '4px' }}>
            Server access token
            <input
              type="password"
              value={app.nas.token ?? ''}
              placeholder="Not set"
              onChange={(e) => (app.nas.token = e.currentTarget.value || null)}
            />
          </label>
        )}
      </div>
    </Dialog>
  );
}
