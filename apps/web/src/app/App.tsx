import { useComputed, useSignal, useSignalEffect, type ReadonlySignal } from '@preact/signals';
import { useEffect, useMemo, useState } from 'preact/hooks';
import { SheetDimensions, type CellValue } from '@cellestra/core';
import {
  CommandMenu,
  Dialog,
  FormatCellsDialog,
  FormulaBar,
  Grid,
  Icon,
  NasBrowser,
  QuickPick,
  SheetTabs,
  type CommandCatalog,
  type GridHost,
  type RecentEntry,
} from '@cellestra/ui';
import {
  AutoSaveSeconds,
  normalizeAutoSave,
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
import { Ribbon } from './Ribbon.tsx';
import { GridContextMenu } from './ContextMenu.tsx';

const TOOLBAR: readonly (string | '|')[] = [
  'file.new',
  'file.open',
  'file.save',
  '|',
  'edit.undo',
  'edit.redo',
  '|',
  'format.bold',
  'format.italic',
  'format.underline',
  'format.strikethrough',
  '|',
  'format.alignLeft',
  'format.alignCenter',
  'format.alignRight',
  'format.wrap',
  'format.merge',
  '|',
  'format.asCurrency',
  'format.asPercent',
  'format.decreaseDecimal',
  'format.increaseDecimal',
  '|',
  'format.borderAll',
  'edit.formatPainter',
  'edit.autoSum',
  'data.sortAsc',
  'data.sortDesc',
  '|',
  'view.freezePanes',
  '|',
  'data.toggleFilter',
  'data.createTable',
  'insert.chartSheet',
  'insert.function',
  '|',
  'insert.rowsAbove',
  'insert.columnsLeft',
  '|',
  'edit.find',
  'edit.clearFormats',
  'format.cells',
];

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
      editText: (row, column) => s.editText(row, column),
      draft: (text) => (draft.value = text),
      commit: (row, column, text) => app.commit(row, column, text),
      clearSelection: () => app.clearContents(),
      edge: (row, column, dr, dc) => app.edge(row, column, dr, dc),
      shortcut: (e, inText) => {
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
          const { row, column } = app.selection.active.peek();
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

  // Shortcuts that fire anywhere (outside the grid too): ⌘K, ⌘S, ⌘O…
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.defaultPrevented || app.dialog.peek() || app.paletteOpen.peek()) return;
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
    addEventListener('keydown', onKey);
    document.addEventListener('copy', onCopy);
    document.addEventListener('cut', onCopy);
    document.addEventListener('paste', onPaste);
    return () => {
      removeEventListener('keydown', onKey);
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
      <div class="cs-main" onPointerUp={() => app.selectionDone()}>
        {s.activeSheet.value.isChartSheet ? (
          <ChartPane app={app} />
        ) : (
          <Grid
            host={host}
            selection={app.selection}
            zoom={app.zoom}
            palette={app.palette}
            showFormulas={app.showFormulas}
            controller={app.grid}
          />
        )}
      </div>
      <GridContextMenu app={app} catalog={catalog} />
      <Footer app={app} />
      <Overlays app={app} catalog={catalog} />
      {dragging.value && <div class="cs-drop">Drop a spreadsheet to open it</div>}
      {app.toast.value && (
        <div
          class="cs-toast"
          role="status"
          style={
            app.toast.value.tone === 'error' ? { background: 'var(--cs-danger)', color: '#fff' } : undefined
          }
        >
          {app.toast.value.text}
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

function Toolbar({
  app,
  catalog,
  titleInline,
}: {
  readonly app: AppController;
  readonly catalog: CommandCatalog;
  /** Show the workbook's name in the toolbar (when there is no title bar of our own). */
  readonly titleInline: boolean;
}) {
  const s = app.session;
  // Read the signals the enabled/active checks depend on, so the toolbar re-renders with them.
  void s.version.value;
  void app.selection.state.value;
  void app.formatPainter.value;
  void app.recents.value;
  return (
    <div class="cs-toolbar" role="toolbar" aria-label="Commands">
      {/* Buttons that don't fit wrap onto a hidden second line: whole buttons drop off the
          end instead of the row overflowing. All of them are in ⌘K and the menus. */}
      <div class="cs-toolbar-items">
        {TOOLBAR.map((id, i) => {
          if (id === '|') return <span key={i} class="cs-sep" />;
          const c = catalog.get(id);
          if (!c) return null;
          const label = c.shortcutText ? `${c.title} (${c.shortcutText})` : c.title;
          return (
            <button
              key={id}
              type="button"
              class="cs-btn"
              title={label}
              aria-label={c.title}
              aria-pressed={c.active ? c.active() : undefined}
              disabled={!catalog.isEnabled(c)}
              onPointerDown={(e) => e.preventDefault() /* keep focus (and the selection) in the grid */}
              onClick={() => void catalog.run(id)}
            >
              <Icon name={c.icon ?? 'square'} />
            </button>
          );
        })}
      </div>
      {titleInline ? (
        <span class="cs-title">
          {s.modified.value ? '• ' : ''}
          {s.title.value}
        </span>
      ) : (
        <span class="cs-toolbar-fill" />
      )}
      <button
        type="button"
        class="cs-btn"
        title="Command Menu (⌘K)"
        aria-label="Command Menu"
        onClick={() => (app.paletteOpen.value = true)}
      >
        <Icon name="command" />
      </button>
    </div>
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
    const r = app.selection.range.value;
    const a = SheetDimensions.columnToLetters(st.column) + (st.row + 1);
    if (r.top === r.bottom && r.left === r.right) return a;
    return `${SheetDimensions.columnToLetters(r.left)}${r.top + 1}:${SheetDimensions.columnToLetters(r.right)}${r.bottom + 1}`;
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
      onCommit={(text) => {
        const { row, column } = app.selection.active.peek();
        if (!app.commit(row, column, text)) return;
        app.selection.select(Math.min(row + 1, 1_048_575), column);
      }}
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
    const timer = setTimeout(() => (summary.value = selectionSummary(app, scope)), 120);
    return () => clearTimeout(timer);
  });
  const inputMessage = useComputed(() => {
    void s.version.value;
    const { row, column } = app.selection.active.value;
    return s.activeSheet.value.findValidation(row, column)?.inputMessage ?? '';
  });
  return (
    <div class="cs-footer">
      <SheetTabs
        names={s.sheetNames.value}
        active={s.activeIndex.value}
        onSelect={(i) => {
          // An open cell edit belongs to the sheet it was typed on.
          if (app.grid.current?.commitEdit() ?? true) s.setActiveSheet(i);
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
      <span style={{ flex: 1 }} />
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
      <span class="cs-status" title="Zoom">
        {Math.round(app.zoom.value * 100)}%
      </span>
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
            else app.deleteSheet(menu.index);
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
    const t = setTimeout(() => addEventListener('pointerdown', close, { once: true }));
    return () => {
      clearTimeout(t);
      removeEventListener('pointerdown', close);
    };
  }, []);
  const item = (a: 'rename' | 'duplicate' | 'delete', label: string, disabled = false) => (
    <li
      role="menuitem"
      aria-disabled={disabled}
      style={disabled ? { opacity: 0.45 } : undefined}
      onPointerDown={(e) => {
        e.stopPropagation();
        if (!disabled) props.onPick(a);
      }}
    >
      {label}
    </li>
  );
  return (
    <ul
      class="cs-palette"
      role="menu"
      style={{
        position: 'fixed',
        left: `${props.x}px`,
        top: `${props.y - 110}px`,
        width: '180px',
        margin: 0,
        padding: '4px',
        listStyle: 'none',
      }}
    >
      {item('rename', 'Rename…')}
      {item('duplicate', 'Duplicate')}
      {item('delete', 'Delete', !props.canDelete)}
    </ul>
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

function Overlays({ app, catalog }: { readonly app: AppController; readonly catalog: CommandCatalog }) {
  const s = app.session;
  const d = app.dialog.value;
  const refocus = () => queueMicrotask(() => app.grid.current?.focus());
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
            Runs entirely on your device: in the browser, from your NAS, or as a desktop app. MIT licensed.
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

function FindDialog({ app }: { readonly app: AppController }) {
  const [query, setQuery] = useState(app.findQuery.peek());
  const [replacement, setReplacement] = useState('');
  const [note, setNote] = useState('');
  const close = () => {
    app.findQuery.value = query;
    app.dialog.value = null;
    queueMicrotask(() => app.grid.current?.focus());
  };
  const next = (back = false) => {
    app.findQuery.value = query;
    setNote(app.findNext(query, back) ? '' : 'No matches.');
  };
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
        <span class="cs-status" style={{ padding: 0 }}>
          {note}
        </span>
      </div>
    </Dialog>
  );
}

function ShortcutsDialog({
  catalog,
  onClose,
}: {
  readonly catalog: CommandCatalog;
  readonly onClose: () => void;
}) {
  const groups = new Map<string, { title: string; keys: string }[]>();
  for (const c of catalog.all) {
    if (!c.shortcutText || c.hidden) continue;
    const list = groups.get(c.group) ?? [];
    list.push({ title: c.title, keys: c.shortcutText });
    groups.set(c.group, list);
  }
  return (
    <Dialog
      title="Keyboard Shortcuts"
      onClose={onClose}
      actions={
        <button type="submit" class="cs-btn primary">
          Done
        </button>
      }
      onSubmit={onClose}
    >
      <div style={{ columns: '2', columnGap: '28px', minWidth: '560px' }}>
        {[...groups].map(([group, items]) => (
          <section key={group} style={{ breakInside: 'avoid', marginBottom: '12px' }}>
            <h3 style={{ margin: '0 0 4px', fontSize: '13px' }}>{group}</h3>
            {items.map((i) => (
              <div
                key={i.title}
                style={{ display: 'flex', gap: '12px', fontSize: '13px', lineHeight: '22px' }}
              >
                <span>{i.title}</span>
                <span class="cs-kbd" style={{ marginLeft: 'auto', opacity: 0.7 }}>
                  {i.keys}
                </span>
              </div>
            ))}
          </section>
        ))}
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
              const v = normalizeAutoSave(Number(e.currentTarget.value));
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
