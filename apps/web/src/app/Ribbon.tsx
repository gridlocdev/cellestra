import type { ComponentChildren } from 'preact';
import { useEffect, useRef, useState } from 'preact/hooks';
import { useSignal } from '@preact/signals';
import { Argb, type ArgbColor, type CellStyle } from '@cellestra/core';
import {
  Icon,
  NUMBER_FORMAT_PRESETS,
  presetIndex,
  ROVING_GROUP,
  rovingItem,
  rovingStep,
  takeFocus,
  type Command,
  type CommandCatalog,
} from '@cellestra/ui';
import type { AppController } from './controller.ts';
import './ribbon.css';

/**
 * The tabbed ribbon (the default layout): a top bar with the workbook's name and command
 * search, tabs, and each tab's commands in labelled groups. Every button runs a catalog
 * command, so the ribbon, the compact toolbar, ⌘K and the menus always agree.
 */

type TabId = 'file' | 'home' | 'insert' | 'formulas' | 'data' | 'review' | 'view';

const TABS: readonly { readonly id: TabId; readonly label: string }[] = [
  { id: 'file', label: 'File' },
  { id: 'home', label: 'Home' },
  { id: 'insert', label: 'Insert' },
  { id: 'formulas', label: 'Formulas' },
  { id: 'data', label: 'Data' },
  { id: 'review', label: 'Review' },
  { id: 'view', label: 'View' },
];

/** Families offered in the font box; the cell's own family is added when it isn't one of them. */
const FONT_FAMILIES = [
  'Helvetica Neue',
  'Arial',
  'Helvetica',
  'Liberation Sans',
  'Calibri',
  'Verdana',
  'Georgia',
  'Times New Roman',
  'Courier New',
  'Menlo',
];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

const TEXT_COLORS = [
  '#000000',
  '#5E6673',
  '#D14343',
  '#E8833A',
  '#D9A400',
  '#2E9E5B',
  '#1F8A9E',
  '#2557D6',
  '#7A4FD6',
  '#C2448F',
];
const FILL_COLORS = [
  '#F2F3F5',
  '#D5D9E0',
  '#FBDADA',
  '#FCE3CF',
  '#FFF2B3',
  '#D6F0DF',
  '#D2EEF2',
  '#DCE6FB',
  '#E7DDFB',
  '#F8DDEB',
];

/** A menu's entries: command ids, with '|' for a divider. */
type MenuItems = readonly string[];

interface Ctx {
  readonly app: AppController;
  readonly catalog: CommandCatalog;
}

export function Ribbon({
  app,
  catalog,
  macTitleBar,
}: {
  readonly app: AppController;
  readonly catalog: CommandCatalog;
  /** macOS window: the top bar is the title bar, with room for the window controls. */
  readonly macTitleBar: boolean;
}) {
  const s = app.session;
  const tab = useSignal<TabId>('home');
  // Re-render with anything the enabled/active checks and the style boxes read.
  void s.version.value;
  void app.selection.state.value;
  void app.formatPainter.value;
  void app.recents.value;
  void app.zoom.value;
  void app.showFormulas.value;
  void app.isDark.value;
  const collapsed = app.ribbonCollapsed.value;
  const ctx: Ctx = { app, catalog };
  const status = s.modified.value ? 'Edited' : s.document.value ? 'Saved' : '';

  return (
    <div class="rb">
      <div class={`rb-top ${macTitleBar ? 'mac' : ''}`}>
        <svg class="rb-logo" viewBox="0 0 24 24" aria-hidden="true">
          <rect x="3" y="3" width="18" height="18" rx="3" />
          <path d="M3 9h18M3 15h18M9 3v18" />
        </svg>
        <span class="rb-name">{s.title.value}</span>
        {status && <span class="rb-status">{status}</span>}
        <div class="rb-fill" />
        <label class="rb-search">
          <Icon name="search" />
          <input
            type="text"
            placeholder="Tell me what you want to do"
            aria-label="Search commands"
            readOnly
            onFocus={(e) => {
              e.currentTarget.blur();
              if (!app.modalOpen()) app.paletteOpen.value = true;
            }}
          />
          <kbd>⌘K</kbd>
        </label>
      </div>

      <div
        class="rb-tabs"
        role="tablist"
        aria-label="Ribbon"
        {...ROVING_GROUP}
        onKeyDown={(e) => {
          const index = TABS.findIndex((t) => t.id === tab.value);
          const next = rovingStep(e, index, TABS.length, 'horizontal');
          if (next === undefined) return;
          e.preventDefault();
          tab.value = TABS[next]!.id;
        }}
      >
        {TABS.map((t) => (
          <button
            key={t.id}
            type="button"
            role="tab"
            class={`rb-tab ${tab.value === t.id && !collapsed ? 'on' : ''}`}
            aria-selected={tab.value === t.id}
            {...rovingItem(tab.value === t.id)}
            onPointerDown={(e) => e.preventDefault()}
            onClick={() => {
              if (collapsed) app.ribbonCollapsed.value = false;
              tab.value = t.id;
            }}
          >
            {t.label}
          </button>
        ))}
      </div>

      {!collapsed && (
        <div class="rb-body" role="tabpanel" aria-label={TABS.find((t) => t.id === tab.value)!.label}>
          {tab.value === 'file' && <FileTab {...ctx} />}
          {tab.value === 'home' && <HomeTab {...ctx} />}
          {tab.value === 'insert' && <InsertTab {...ctx} />}
          {tab.value === 'formulas' && <FormulasTab {...ctx} />}
          {tab.value === 'data' && <DataTab {...ctx} />}
          {tab.value === 'review' && <ReviewTab {...ctx} />}
          {tab.value === 'view' && <ViewTab {...ctx} />}
          <div class="rb-fill" />
          <button
            type="button"
            class="rb-tb rb-collapse"
            title="Collapse Ribbon"
            aria-label="Collapse Ribbon"
            onPointerDown={(e) => e.preventDefault()}
            onClick={() => (app.ribbonCollapsed.value = true)}
          >
            <Icon name="chevron-up" />
          </button>
        </div>
      )}
    </div>
  );
}

// --- tabs ------------------------------------------------------------------------------------

function HomeTab(ctx: Ctx) {
  const { app } = ctx;
  const st = app.activeStyle;
  // A chart sheet has no cells to style (the command buttons disable themselves).
  const off = app.session.activeSheet.value.isChartSheet;
  return (
    <>
      <Group label="Clipboard">
        <div class="rb-row">
          <Big {...ctx} id="edit.paste" label="Paste" menu={['edit.paste', 'edit.pasteSpecial']} />
          <Col>
            <Small {...ctx} id="edit.cut" label="Cut" />
            <Small {...ctx} id="edit.copy" label="Copy" />
            <Small {...ctx} id="edit.formatPainter" label="Format painter" />
          </Col>
        </div>
      </Group>

      <Group label="Font">
        <Stack>
          <div class="rb-row">
            <FontFamilyBox app={app} style={st} disabled={off} />
            <FontSizeBox app={app} style={st} disabled={off} />
          </div>
          <div class="rb-row tight">
            <Tool {...ctx} id="format.bold" glyph={<b class="rb-g">B</b>} />
            <Tool {...ctx} id="format.italic" glyph={<i class="rb-g serif">I</i>} />
            <Tool {...ctx} id="format.underline" glyph={<u class="rb-g">U</u>} />
            <Tool {...ctx} id="format.strikethrough" glyph={<s class="rb-g">S</s>} />
            <Sep />
            <ColorButton
              app={app}
              disabled={off}
              kind="text"
              current={st.font.color}
              swatches={TEXT_COLORS}
              resetLabel="Automatic"
            />
            <ColorButton
              app={app}
              disabled={off}
              kind="fill"
              current={st.fill.background}
              swatches={FILL_COLORS}
              resetLabel="No fill"
            />
            <MenuTool
              {...ctx}
              icon="grid-2x2"
              label="Borders"
              items={[
                'format.borderAll',
                'format.borderOutline',
                '|',
                'format.borderTop',
                'format.borderBottom',
                'format.borderLeft',
                'format.borderRight',
                '|',
                'format.borderNone',
              ]}
            />
          </div>
        </Stack>
      </Group>

      <Group label="Alignment">
        <Stack>
          <div class="rb-row tight">
            <Tool {...ctx} id="format.alignTop" />
            <Tool {...ctx} id="format.alignMiddle" />
            <Tool {...ctx} id="format.alignBottom" />
            <Sep />
            <Small {...ctx} id="format.wrap" label="Wrap text" tall />
          </div>
          <div class="rb-row tight">
            <Tool {...ctx} id="format.alignLeft" />
            <Tool {...ctx} id="format.alignCenter" />
            <Tool {...ctx} id="format.alignRight" />
            <Sep />
            <Small {...ctx} id="format.merge" label="Merge" tall />
          </div>
        </Stack>
      </Group>

      <Group label="Number">
        <Stack>
          <NumberFormatBox app={app} style={st} disabled={off} />
          <div class="rb-row tight">
            <Tool {...ctx} id="format.asCurrency" glyph={<span class="rb-g">$</span>} />
            <Tool {...ctx} id="format.asPercent" glyph={<span class="rb-g">%</span>} />
            <Tool {...ctx} id="format.asNumber" glyph={<span class="rb-g small">000</span>} />
            <Tool {...ctx} id="format.decreaseDecimal" glyph={<span class="rb-g small">.0</span>} />
            <Tool {...ctx} id="format.increaseDecimal" glyph={<span class="rb-g small">.00</span>} />
          </div>
        </Stack>
      </Group>

      <Group label="Cells">
        <Col>
          <MenuSmall
            {...ctx}
            icon="plus"
            label="Insert"
            items={['insert.rowsAbove', 'insert.columnsLeft', '|', 'insert.sheet']}
          />
          <MenuSmall
            {...ctx}
            icon="minus"
            label="Delete"
            items={['edit.deleteRows', 'edit.deleteColumns', '|', 'sheet.delete']}
          />
          <MenuSmall
            {...ctx}
            icon="table"
            label="Format"
            items={[
              'format.cells',
              'format.conditional',
              '|',
              'format.columnWidth',
              'format.rowHeight',
              'data.autofitColumns',
              '|',
              'format.hideRows',
              'format.unhideRows',
              'format.hideColumns',
              'format.unhideColumns',
              '|',
              'sheet.rename',
            ]}
          />
        </Col>
      </Group>

      <Group label="Editing">
        <Col>
          <Small
            {...ctx}
            id="edit.autoSum"
            label="AutoSum"
            menu={['edit.autoSum', 'insert.function', '|', 'fill.down', 'fill.right']}
          />
          <MenuSmall
            {...ctx}
            icon="filter"
            label="Sort & filter"
            items={[
              'data.sortAsc',
              'data.sortDesc',
              'data.customSort',
              '|',
              'data.toggleFilter',
              'data.clearFilter',
              'data.sortAndFilter',
            ]}
          />
          <Small {...ctx} id="edit.find" label="Find & select" />
        </Col>
      </Group>
    </>
  );
}

function FileTab(ctx: Ctx) {
  return (
    <>
      <Group label="Workbook">
        <div class="rb-row">
          <Big {...ctx} id="file.new" label="New" />
          <Big {...ctx} id="file.open" label="Open" />
          <Big {...ctx} id="file.save" label="Save" />
          <Col>
            <Small {...ctx} id="file.saveAs" label="Save as…" />
            <Small {...ctx} id="file.download" label="Download" />
            <Small {...ctx} id="file.revert" label="Revert" />
          </Col>
          <Col>
            <Small {...ctx} id="file.close" label="Close" />
          </Col>
        </div>
      </Group>
      <Group label="Recent">
        <Col>
          <Small {...ctx} id="file.recent" label="Recent files" />
          <Small {...ctx} id="file.openLocal" label="Open from computer" />
        </Col>
      </Group>
      <Group label="CSV">
        <Col>
          <Small {...ctx} id="file.importCsv" label="Import CSV" />
          <Small {...ctx} id="file.exportCsv" label="Export CSV" />
        </Col>
      </Group>
      <Group label="Print">
        <div class="rb-row">
          <Big {...ctx} id="file.print" label="Print" />
          <Col>
            <Small {...ctx} id="file.pageSetup" label="Page setup" />
            <Small {...ctx} id="file.printSelection" label="Print selection" />
            <Small {...ctx} id="file.exportPdf" label="Export PDF" />
          </Col>
        </div>
      </Group>
      <Group label="Options">
        <Big {...ctx} id="app.preferences" label="Settings" />
      </Group>
    </>
  );
}

function InsertTab(ctx: Ctx) {
  return (
    <>
      <Group label="Sheets">
        <div class="rb-row">
          <Big {...ctx} id="insert.sheet" label="Sheet" />
          <Big {...ctx} id="insert.chartSheet" label="Chart" />
        </div>
      </Group>
      <Group label="Cells">
        <Col>
          <Small {...ctx} id="insert.rowsAbove" label="Rows above" />
          <Small {...ctx} id="insert.columnsLeft" label="Columns left" />
        </Col>
      </Group>
      <Group label="Formula">
        <Big {...ctx} id="insert.function" label="Function" />
      </Group>
      <Group label="Controls">
        <Big {...ctx} id="insert.checkbox" label="Checkbox" />
      </Group>
      <Group label="Links & names">
        <Col>
          <Small {...ctx} id="insert.hyperlink" label="Hyperlink" />
          <Small {...ctx} id="insert.namedRange" label="Named range" />
        </Col>
      </Group>
      <Group label="Date & time">
        <Col>
          <Small {...ctx} id="insert.date" label="Today's date" />
          <Small {...ctx} id="insert.time" label="Current time" />
        </Col>
      </Group>
    </>
  );
}

function FormulasTab(ctx: Ctx) {
  return (
    <>
      <Group label="Function library">
        <div class="rb-row">
          <Big {...ctx} id="insert.function" label="Insert function" />
          <Col>
            <Small {...ctx} id="edit.autoSum" label="AutoSum" />
            <Small {...ctx} id="fill.down" label="Fill down" />
            <Small {...ctx} id="fill.right" label="Fill right" />
          </Col>
        </div>
      </Group>
      <Group label="Names">
        <Big {...ctx} id="insert.namedRange" label="Named ranges" />
      </Group>
      <Group label="Calculation">
        <div class="rb-row">
          <Big {...ctx} id="formula.recalc" label="Recalculate" />
          <Col>
            <Small {...ctx} id="formula.manualCalc" label="Manual calculation" />
            <Small {...ctx} id="view.showFormulas" label="Show formulas" />
          </Col>
        </div>
      </Group>
    </>
  );
}

function DataTab(ctx: Ctx) {
  return (
    <>
      <Group label="Sort & filter">
        <div class="rb-row">
          <Col>
            <Small {...ctx} id="data.sortAsc" label="Sort A → Z" />
            <Small {...ctx} id="data.sortDesc" label="Sort Z → A" />
            <Small {...ctx} id="data.customSort" label="Custom sort…" />
          </Col>
          <Big {...ctx} id="data.toggleFilter" label="Filter" />
          <Col>
            <Small {...ctx} id="data.clearFilter" label="Clear filter" />
            <Small {...ctx} id="data.sortAndFilter" label="Sort & filter…" />
          </Col>
        </div>
      </Group>
      <Group label="Table">
        <div class="rb-row">
          <Big {...ctx} id="data.createTable" label="Table" />
          <Col>
            <TableNameBox app={ctx.app} />
            <Small {...ctx} id="data.bandedRows" label="Banded rows" />
            <Small {...ctx} id="data.headerColumn" label="Header column" />
            <Small {...ctx} id="data.removeTable" label="Remove table" />
          </Col>
        </div>
      </Group>
      <Group label="Data tools">
        <Col>
          <Small {...ctx} id="data.validation" label="Validation…" />
          <Small {...ctx} id="data.removeDuplicates" label="Remove duplicates" />
          <Small {...ctx} id="data.textToColumns" label="Text to columns" />
        </Col>
      </Group>
      <Group label="Columns">
        <Big {...ctx} id="data.autofitColumns" label="Autofit" />
      </Group>
    </>
  );
}

function ReviewTab(ctx: Ctx) {
  return (
    <>
      <Group label="Find">
        <div class="rb-row">
          <Big {...ctx} id="edit.find" label="Find" />
          <Col>
            <Small {...ctx} id="edit.findReplace" label="Replace" />
            <Small {...ctx} id="edit.findNext" label="Find next" />
          </Col>
        </div>
      </Group>
      <Group label="Clean up">
        <div class="rb-row">
          <Big {...ctx} id="format.cleanUpSheet" label="Clean up sheet" />
          <Col>
            <Small {...ctx} id="format.cleanUpWorkbook" label="Whole workbook" />
            <Small {...ctx} id="edit.clearFormats" label="Clear formats" />
            <Small {...ctx} id="edit.clearAll" label="Clear all" />
          </Col>
        </div>
      </Group>
    </>
  );
}

function ViewTab(ctx: Ctx) {
  return (
    <>
      <Group label="Show">
        <div class="rb-row">
          <Big {...ctx} id="view.freezePanes" label="Freeze panes" />
          <Col>
            <Small {...ctx} id="view.showFormulas" label="Formulas" />
            <Small {...ctx} id="view.goToSheet" label="Go to sheet…" />
          </Col>
        </div>
      </Group>
      <Group label="Zoom">
        <Col>
          <Small {...ctx} id="view.zoomIn" label="Zoom in" />
          <Small {...ctx} id="view.zoomOut" label="Zoom out" />
          <Small {...ctx} id="view.zoomReset" label="Actual size" />
        </Col>
      </Group>
      <Group label="Window">
        <Col>
          <Small {...ctx} id="view.theme" label="Dark mode" />
          <Small {...ctx} id="view.toggleLayout" label="Ribbon layout" />
          <Small {...ctx} id="app.preferences" label="Settings…" />
        </Col>
      </Group>
      <Group label="Help">
        <Col>
          <Small {...ctx} id="view.commandMenu" label="Command menu" />
          <Small {...ctx} id="help.shortcuts" label="Keyboard shortcuts" />
          <Small {...ctx} id="help.about" label="About" />
        </Col>
      </Group>
    </>
  );
}

// --- building blocks ---------------------------------------------------------------------------

function Group({ label, children }: { readonly label: string; readonly children: ComponentChildren }) {
  return (
    <div class="rb-grp" role="group" aria-label={label}>
      <div class="rb-grp-body">{children}</div>
      <div class="rb-grp-label">{label}</div>
    </div>
  );
}

const Col = ({ children }: { readonly children: ComponentChildren }) => <div class="rb-col">{children}</div>;
const Stack = ({ children }: { readonly children: ComponentChildren }) => (
  <div class="rb-stack">{children}</div>
);
const Sep = () => <div class="rb-sep" />;

/** The catalog command, or null when this build doesn't have it (or hides it). */
function usable(catalog: CommandCatalog, id: string): Command | null {
  const c = catalog.get(id);
  return c && !c.hidden ? c : null;
}

function tip(c: Command): string {
  return c.shortcutText ? `${c.title} (${c.shortcutText})` : c.title;
}

/** Common button props: keep focus (and the selection) in the grid, run the command. */
function commandProps({ app, catalog }: Ctx, c: Command) {
  return {
    type: 'button' as const,
    title: tip(c),
    'aria-label': c.title,
    'aria-pressed': c.active ? c.active() : undefined,
    disabled: !catalog.isEnabled(c),
    onPointerDown: (e: PointerEvent) => e.preventDefault(),
    onClick: () => {
      void catalog.run(c.id);
      app.grid.current?.focus();
    },
  };
}

function Big(props: Ctx & { readonly id: string; readonly label: string; readonly menu?: MenuItems }) {
  const c = usable(props.catalog, props.id);
  if (!c) return null;
  if (!props.menu) {
    return (
      <button class="rb-big" {...commandProps(props, c)}>
        <Icon name={c.icon ?? 'square'} />
        <span>{props.label}</span>
      </button>
    );
  }
  // Split: the icon runs the command, the label opens the menu.
  return (
    <div class="rb-big split">
      <button class="rb-big-top" {...commandProps(props, c)}>
        <Icon name={c.icon ?? 'square'} />
      </button>
      <MenuAnchor {...props} items={props.menu} buttonClass="rb-big-bottom" label={`${props.label} options`}>
        {props.label}
        <Icon name="chevron-down" />
      </MenuAnchor>
    </div>
  );
}

function Small(
  props: Ctx & {
    readonly id: string;
    readonly label: string;
    readonly menu?: MenuItems;
    /** Full button height (32px), for a labelled button in a row of icon buttons. */
    readonly tall?: boolean;
  },
) {
  const c = usable(props.catalog, props.id);
  if (!c) return null;
  const cls = `rb-tb rb-sm ${props.tall ? 'tall' : ''}`;
  if (!props.menu) {
    return (
      <button class={cls} {...commandProps(props, c)}>
        <Icon name={c.icon ?? 'square'} />
        {props.label}
      </button>
    );
  }
  return (
    <div class="rb-split">
      <button class={cls} {...commandProps(props, c)}>
        <Icon name={c.icon ?? 'square'} />
        {props.label}
      </button>
      <MenuAnchor
        {...props}
        items={props.menu}
        buttonClass="rb-tb rb-sm rb-chev"
        label={`${props.label} options`}
      >
        <Icon name="chevron-down" />
      </MenuAnchor>
    </div>
  );
}

function Tool(props: Ctx & { readonly id: string; readonly glyph?: ComponentChildren }) {
  const c = usable(props.catalog, props.id);
  if (!c) return null;
  return (
    <button class="rb-tb" {...commandProps(props, c)}>
      {props.glyph ?? <Icon name={c.icon ?? 'square'} />}
    </button>
  );
}

/** A labelled small button whose whole face opens a menu. */
function MenuSmall(
  props: Ctx & { readonly icon: string; readonly label: string; readonly items: MenuItems },
) {
  return (
    <MenuAnchor {...props} buttonClass="rb-tb rb-sm" label={props.label}>
      <Icon name={props.icon} />
      {props.label}
      <Icon name="chevron-down" />
    </MenuAnchor>
  );
}

/** An icon button with a chevron that opens a menu. */
function MenuTool(props: Ctx & { readonly icon: string; readonly label: string; readonly items: MenuItems }) {
  return (
    <MenuAnchor {...props} buttonClass="rb-tb" label={props.label}>
      <Icon name={props.icon} />
      <Icon name="chevron-down" />
    </MenuAnchor>
  );
}

/** A button that opens a drop-down list of commands under itself. */
function MenuAnchor(
  props: Ctx & {
    readonly items: MenuItems;
    readonly buttonClass: string;
    readonly label: string;
    readonly children: ComponentChildren;
  },
) {
  const [open, setOpen] = useState(false);
  // Opened from the keyboard, the menu takes it, on its current item; from the mouse, the grid keeps it.
  const [byKeyboard, setByKeyboard] = useState(false);
  const commands = props.items.map((id) => (id === '|' ? '|' : usable(props.catalog, id)));
  const pickable = commands.flatMap((c) => (c && c !== '|' && props.catalog.isEnabled(c) ? [c.id] : []));
  const [current, setCurrent] = useState(0);
  return (
    <Popover
      open={open}
      onClose={() => setOpen(false)}
      anchor={
        <button
          type="button"
          class={props.buttonClass}
          title={props.label}
          aria-label={props.label}
          aria-haspopup="menu"
          aria-expanded={open}
          onPointerDown={(e) => e.preventDefault()}
          onClick={(e) => {
            setByKeyboard(e.detail === 0);
            setCurrent(0);
            setOpen(!open);
          }}
        >
          {props.children}
        </button>
      }
    >
      <div
        class="rb-menu"
        role="menu"
        aria-label={props.label}
        {...ROVING_GROUP}
        onKeyDown={(e) => {
          const next = rovingStep(e, current, pickable.length, 'vertical');
          if (next === undefined) return;
          e.preventDefault();
          setCurrent(next);
        }}
      >
        {commands.map((c, i) =>
          c === '|' ? (
            <div key={i} class="rb-menu-sep" role="separator" />
          ) : c ? (
            <button
              key={c.id}
              type="button"
              role="menuitem"
              class="rb-menu-item"
              disabled={!props.catalog.isEnabled(c)}
              {...rovingItem(pickable[current] === c.id, byKeyboard ? takeFocus : undefined)}
              onPointerDown={(e) => e.preventDefault()}
              onClick={() => {
                setOpen(false);
                void props.catalog.run(c.id);
                props.app.grid.current?.focus();
              }}
            >
              <Icon name={c.icon ?? 'square'} />
              <span class="rb-menu-title">{c.title}</span>
              {c.shortcutText && <span class="rb-menu-key">{c.shortcutText}</span>}
            </button>
          ) : null,
        )}
      </div>
    </Popover>
  );
}

/** Something anchored under a button; closes on a click elsewhere or Escape. */
function Popover(props: {
  readonly open: boolean;
  readonly onClose: () => void;
  readonly anchor: ComponentChildren;
  readonly children: ComponentChildren;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const { open, onClose } = props;
  useEffect(() => {
    if (!open) return;
    const away = (e: PointerEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose();
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        // Focus inside the panel goes back to the button that opened it, not to the page.
        const inside = ref.current?.querySelector('.rb-pop-panel')?.contains(document.activeElement);
        onClose();
        if (inside) ref.current?.querySelector<HTMLElement>('[aria-haspopup]')?.focus();
      }
    };
    addEventListener('pointerdown', away, true);
    addEventListener('keydown', key, true);
    return () => {
      removeEventListener('pointerdown', away, true);
      removeEventListener('keydown', key, true);
    };
  }, [open, onClose]);
  // The ribbon body scrolls sideways (and so clips), so the panel is fixed, placed under the anchor.
  const at = open ? ref.current?.getBoundingClientRect() : undefined;
  const left = at ? Math.max(4, Math.min(at.left, innerWidth - 200)) : 0;
  return (
    <div class="rb-pop" ref={ref}>
      {props.anchor}
      {open && at && (
        <div class="rb-pop-panel" style={{ top: `${at.bottom}px`, left: `${left}px` }}>
          {props.children}
        </div>
      )}
    </div>
  );
}

// --- style boxes -------------------------------------------------------------------------------

function refocus(app: AppController) {
  queueMicrotask(() => app.grid.current?.focus());
}

interface BoxProps {
  readonly app: AppController;
  readonly style: CellStyle;
  readonly disabled: boolean;
}

function FontFamilyBox({ app, style, disabled }: BoxProps) {
  const family = style.font.family;
  const families = FONT_FAMILIES.includes(family) ? FONT_FAMILIES : [family, ...FONT_FAMILIES];
  return (
    <select
      class="rb-dd"
      style={{ width: '150px' }}
      aria-label="Font"
      title="Font"
      disabled={disabled}
      value={family}
      onChange={(e) => {
        const next = e.currentTarget.value;
        app.style('Font', (st) => ({ ...st, font: { ...st.font, family: next } }));
        refocus(app);
      }}
    >
      {families.map((f) => (
        <option key={f} value={f}>
          {f}
        </option>
      ))}
    </select>
  );
}

/**
 * The font size: type a size, or open the list, where hovering a size previews it on the
 * selection before anything is applied.
 */
function FontSizeBox({ app, style, disabled }: BoxProps) {
  const size = style.font.sizePoints;
  const sizes = FONT_SIZES.includes(size) ? FONT_SIZES : [...FONT_SIZES, size].sort((a, b) => a - b);
  const [open, setOpen] = useState(false);
  const [byKeyboard, setByKeyboard] = useState(false);
  // The highlighted size in the list: the current one to start with; ↑ ↓ Home End move and preview it.
  const [at, setAt] = useState(0);
  const [text, setText] = useState(String(size));
  useEffect(() => setText(String(size)), [size]);
  const resize =
    (points: number) =>
    (st: CellStyle): CellStyle => ({ ...st, font: { ...st.font, sizePoints: points } });
  const apply = (points: number) => {
    setOpen(false);
    if (Number.isFinite(points) && points >= 1 && points <= 409 && points !== size)
      app.style('Font Size', resize(points));
    else app.previewStyle(null);
    refocus(app);
  };
  const close = () => {
    setOpen(false);
    app.previewStyle(null);
  };
  return (
    <Popover
      open={open}
      onClose={close}
      anchor={
        <span class="rb-size">
          <input
            class="rb-dd rb-size-input cs-focus-inset"
            aria-label="Font size"
            title="Font size"
            disabled={disabled}
            inputMode="decimal"
            value={text}
            onInput={(e) => setText(e.currentTarget.value)}
            onFocus={(e) => e.currentTarget.select()}
            onBlur={() => setText(String(size))}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                apply(Number(text));
              } else if (e.key === 'Escape') {
                e.preventDefault();
                setText(String(size));
                refocus(app);
              } else if (e.key === 'ArrowDown') {
                e.preventDefault();
                setOpen(true);
              }
            }}
          />
          <button
            type="button"
            class="rb-size-open"
            aria-label="Font sizes"
            aria-haspopup="listbox"
            disabled={disabled}
            aria-expanded={open}
            onPointerDown={(e) => e.preventDefault()}
            onClick={(e) => {
              if (open) return close();
              setByKeyboard(e.detail === 0);
              setAt(Math.max(0, sizes.indexOf(size)));
              setOpen(true);
            }}
          >
            <Icon name="chevron-down" />
          </button>
        </span>
      }
    >
      <div
        class="rb-menu rb-size-list"
        role="listbox"
        aria-label="Font sizes"
        {...ROVING_GROUP}
        onKeyDown={(e) => {
          const next = rovingStep(e, at, sizes.length, 'vertical');
          if (next === undefined) return;
          e.preventDefault();
          setAt(next);
          app.previewStyle(resize(sizes[next]!));
        }}
      >
        {sizes.map((n, i) => (
          <button
            key={n}
            type="button"
            role="option"
            aria-selected={n === size}
            {...rovingItem(i === at, byKeyboard ? takeFocus : undefined)}
            class={`rb-menu-item${n === size ? ' on' : ''}`}
            onPointerDown={(e) => e.preventDefault()}
            onPointerEnter={() => app.previewStyle(resize(n))}
            onPointerLeave={() => app.previewStyle(null)}
            onClick={() => apply(n)}
          >
            {n}
          </button>
        ))}
      </div>
    </Popover>
  );
}

/**
 * The active table's name: type a new one and press Enter (or leave the box) to rename it; Escape
 * puts the name back. Blank and disabled when the cursor isn't in a table.
 */
function TableNameBox({ app }: { readonly app: AppController }) {
  const table = app.sheet.isChartSheet ? null : app.activeTable;
  const name = table?.name ?? '';
  const [text, setText] = useState(name);
  useEffect(() => setText(name), [name]);
  const commit = () => {
    if (table && text.trim() !== name && !app.renameTable(text)) setText(name);
  };
  return (
    <label class="rb-table-name" title="The table’s name. Charts and the Name Box find the table by it.">
      <span>Name</span>
      <input
        class="rb-dd"
        aria-label="Table name"
        disabled={!table}
        placeholder={table ? '' : 'No table'}
        value={text}
        onInput={(e) => setText(e.currentTarget.value)}
        onFocus={(e) => e.currentTarget.select()}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            commit();
            refocus(app);
          } else if (e.key === 'Escape') {
            e.preventDefault();
            setText(name);
            refocus(app);
          }
        }}
      />
    </label>
  );
}

function NumberFormatBox({ app, style, disabled }: BoxProps) {
  const index = presetIndex(style.numberFormat.code);
  return (
    <select
      class="rb-dd"
      style={{ width: '150px' }}
      aria-label="Number format"
      title="Number format"
      disabled={disabled}
      value={String(index)}
      onChange={(e) => {
        const preset = NUMBER_FORMAT_PRESETS[Number(e.currentTarget.value)];
        if (!preset) return;
        app.style(`${preset.label} Format`, (st) => ({ ...st, numberFormat: { code: preset.code } }));
        refocus(app);
      }}
    >
      {index < 0 && (
        <option value="-1" disabled>
          Custom
        </option>
      )}
      {NUMBER_FORMAT_PRESETS.map((p, i) => (
        <option key={p.code} value={String(i)}>
          {p.label}
        </option>
      ))}
    </select>
  );
}

/** Text colour or fill: the icon shows the current colour; the swatches apply one. */
function ColorButton(props: {
  readonly app: AppController;
  readonly kind: 'text' | 'fill';
  readonly current: ArgbColor | null;
  readonly swatches: readonly string[];
  readonly resetLabel: string;
  readonly disabled: boolean;
}) {
  const { app, kind } = props;
  const [open, setOpen] = useState(false);
  const label = kind === 'text' ? 'Text color' : 'Fill color';
  const change =
    (color: ArgbColor | null) =>
    (st: CellStyle): CellStyle =>
      kind === 'text'
        ? { ...st, font: { ...st.font, color: color ?? Argb.black } }
        : { ...st, fill: { background: color } };
  const apply = (color: ArgbColor | null) => {
    setOpen(false);
    app.style(kind === 'text' ? 'Text Color' : 'Fill Color', change(color));
    refocus(app);
  };
  // Hovering a swatch shows it on the selection; leaving (or closing) puts the real colours back.
  const preview = (color: ArgbColor | null | undefined) =>
    app.previewStyle(color === undefined ? null : change(color));
  const close = () => {
    setOpen(false);
    preview(undefined);
  };
  const bar =
    props.current === null || Argb.isTransparent(props.current)
      ? 'transparent'
      : kind === 'text' && props.current === Argb.black
        ? 'currentColor'
        : Argb.toHex(props.current);
  return (
    <Popover
      open={open}
      onClose={close}
      anchor={
        <button
          type="button"
          class="rb-tb"
          title={label}
          aria-label={label}
          aria-haspopup="dialog"
          aria-expanded={open}
          disabled={props.disabled}
          onPointerDown={(e) => e.preventDefault()}
          onClick={() => setOpen(!open)}
        >
          <span class="rb-color">
            {kind === 'text' ? (
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <path d="m7 15 5-11 5 11" />
                <path d="M9 11h6" />
              </svg>
            ) : (
              <Icon name="paint-bucket" />
            )}
            <span class="rb-color-bar" style={{ background: bar }} />
          </span>
        </button>
      }
    >
      <div class="rb-swatches" role="dialog" aria-label={label}>
        <button
          type="button"
          class="rb-swatch-reset"
          onPointerDown={(e) => e.preventDefault()}
          onClick={() => apply(null)}
          onPointerEnter={() => preview(null)}
          onPointerLeave={() => preview(undefined)}
        >
          {props.resetLabel}
        </button>
        <div class="rb-swatch-grid">
          {props.swatches.map((hex) => (
            <button
              key={hex}
              type="button"
              class="rb-swatch"
              title={hex}
              aria-label={hex}
              style={{ background: hex }}
              onPointerDown={(e) => e.preventDefault()}
              onClick={() => apply(Argb.fromHex(hex))}
              onPointerEnter={() => preview(Argb.fromHex(hex))}
              onPointerLeave={() => preview(undefined)}
            />
          ))}
        </div>
      </div>
    </Popover>
  );
}
