import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { PageSetups, type PageScale, type PageSetup } from '../../model/pageSetup.ts';
import { XmlElement, xmlNamespace, type XmlDocument, type XmlName } from '../../xml/index.ts';
import { formatUpTo3 } from './odsContentShared.ts';
import { parseLength } from './odsStyleIndex.ts';
import { Fo, Office, Style, Table } from './odsXmlNames.ts';

/** Names a save gives the page layouts, master pages and table styles it writes. */
const LAYOUT_PREFIX = 'CellestraPL';
const MASTER_PREFIX = 'CellestraPage';
const TABLE_STYLE_PREFIX = 'CellestraTa';

/** What each sheet's page looked like in the file, so a save rewrites only the pages that changed. */
interface ReadPage {
  readonly setup: PageSetup;
  readonly master: string;
  readonly tableStyle: string | null;
}
const readPages = new WeakMap<Worksheet, ReadPage>();

/** The page styles a save writes into content.xml. */
export interface PagePlan {
  /** New automatic table styles, each naming its master page. */
  readonly tableStyles: readonly XmlElement[];
  /** The table style each regenerated sheet uses. */
  readonly tableStyleNames: ReadonlyMap<Worksheet, string>;
  /** Whether styles.xml changed. */
  readonly stylesChanged: boolean;
}

export const NO_PAGE_PLAN: PagePlan = { tableStyles: [], tableStyleNames: new Map(), stylesChanged: false };

/** Whether a preserved automatic style is one an earlier save wrote for a page, now replaced. */
export const isGeneratedTableStyle = (style: XmlElement): boolean =>
  style.attr(Style('name'))?.startsWith(TABLE_STYLE_PREFIX) ?? false;

export const OdsPageSetup = {
  /**
   * Give each sheet the page its table style's master page uses (`style:master-page-name`, else
   * the "Default" master page, as LibreOffice does).
   */
  read(workbook: Workbook, content: XmlElement, styles: XmlElement | null): void {
    if (!styles) return;
    const masters = new Map<string, string>();
    for (const m of styles.element(Office('master-styles'))?.elements(Style('master-page')) ?? []) {
      const name = m.attr(Style('name'));
      const layout = m.attr(Style('page-layout-name'));
      if (name !== null && layout !== null) masters.set(name, layout);
    }
    const layouts = new Map<string, XmlElement>();
    for (const l of styles.element(Office('automatic-styles'))?.elements(Style('page-layout')) ?? []) {
      const name = l.attr(Style('name'));
      if (name !== null) layouts.set(name, l);
    }
    const tableMasters = new Map<string, string>();
    for (const s of content.element(Office('automatic-styles'))?.elements(Style('style')) ?? []) {
      const name = s.attr(Style('name'));
      const master = s.attr(Style('master-page-name'));
      if (s.attr(Style('family')) === 'table' && name !== null && master) tableMasters.set(name, master);
    }
    const spreadsheet = content.element(Office('body'))?.element(Office('spreadsheet'));
    for (const table of spreadsheet?.elements(Table('table')) ?? []) {
      const sheet = workbook.findSheet(table.attr(Table('name')) ?? '');
      if (!sheet) continue;
      const tableStyle = table.attr(Table('style-name'));
      const master = (tableStyle !== null ? tableMasters.get(tableStyle) : undefined) ?? 'Default';
      const layout = layouts.get(masters.get(master) ?? '');
      if (!layout) continue;
      const setup = readLayout(layout);
      sheet.pageSetup = setup;
      readPages.set(sheet, { setup, master, tableStyle });
    }
  },

  /**
   * Write a page layout and master page into styles.xml for each sheet whose page changed (or was
   * written by an earlier save), copying the page it had so its headers and footers stay. Pages an
   * earlier save wrote are dropped first. `preservedStyles` are content.xml's kept automatic styles.
   */
  write(workbook: Workbook, styles: XmlDocument, preservedStyles: readonly XmlElement[]): PagePlan {
    const root = styles.root;
    const masterStyles = root.element(Office('master-styles'));
    const automatic = root.element(Office('automatic-styles'));
    const masters = new Map((masterStyles?.elements(Style('master-page')) ?? []).map((m) => [nameOf(m), m]));
    const layouts = new Map((automatic?.elements(Style('page-layout')) ?? []).map((l) => [nameOf(l), l]));
    const tableStyles = new Map(preservedStyles.map((s) => [nameOf(s), s]));

    let stylesChanged = false;
    for (const [name, m] of masters)
      if (name.startsWith(MASTER_PREFIX)) stylesChanged = masterStyles!.remove(m) || stylesChanged;
    for (const [name, l] of layouts)
      if (name.startsWith(LAYOUT_PREFIX)) stylesChanged = automatic!.remove(l) || stylesChanged;

    const used = new Set<string>();
    for (const e of root.descendants()) used.add(nameOf(e));
    for (const s of preservedStyles) if (!isGeneratedTableStyle(s)) used.add(nameOf(s));
    const fresh = (prefix: string): string => {
      let i = 0;
      while (used.has(`${prefix}${i}`)) i++;
      used.add(`${prefix}${i}`);
      return `${prefix}${i}`;
    };

    const newTableStyles: XmlElement[] = [];
    const tableStyleNames = new Map<Worksheet, string>();
    for (const sheet of workbook.worksheets) {
      const before = readPages.get(sheet);
      const generated = before?.master.startsWith(MASTER_PREFIX) ?? false;
      // A page an earlier save wrote is always written again: its old styles were dropped.
      const setup = sheet.pageSetup ?? (generated ? PageSetups.default : null);
      if (setup === null) continue;
      if (before && !generated && PageSetups.equal(before.setup, setup)) continue;

      const sourceMaster = before ? masters.get(before.master) : (masters.get('Default') ?? null);
      const sourceLayout = sourceMaster
        ? layouts.get(sourceMaster.attr(Style('page-layout-name')) ?? '')
        : null;
      const layoutName = fresh(LAYOUT_PREFIX);
      const layout = sourceLayout?.clone() ?? new XmlElement(Style('page-layout'));
      layout.setAttr(Style('name'), layoutName);
      writeLayout(layout, setup);
      const masterName = fresh(MASTER_PREFIX);
      const master = sourceMaster?.clone() ?? new XmlElement(Style('master-page'));
      master.setAttr(Style('name'), masterName).setAttr(Style('display-name'), null);
      master.setAttr(Style('page-layout-name'), layoutName);
      child(root, Office('automatic-styles'), [Office('styles'), Office('font-face-decls')]).add(layout);
      child(root, Office('master-styles'), [
        Office('automatic-styles'),
        Office('styles'),
        Office('font-face-decls'),
      ]).add(master);
      stylesChanged = true;

      const sourceTable =
        before && before.tableStyle !== null ? tableStyles.get(before.tableStyle) : undefined;
      const tableStyleName = fresh(TABLE_STYLE_PREFIX);
      const tableStyle =
        sourceTable?.clone() ??
        new XmlElement(
          Style('style'),
          new XmlElement(Style('table-properties')).setAttr(Table('display'), 'true'),
        );
      tableStyle
        .setAttr(Style('name'), tableStyleName)
        .setAttr(Style('family'), 'table')
        .setAttr(Style('master-page-name'), masterName);
      newTableStyles.push(tableStyle);
      tableStyleNames.set(sheet, tableStyleName);
    }
    if (stylesChanged) root.declare(Style.prefix, Style.uri).declare(Fo.prefix, Fo.uri);
    return { tableStyles: newTableStyles, tableStyleNames, stylesChanged };
  },
} as const;

const nameOf = (e: XmlElement): string => e.attr(Style('name')) ?? '';

/** `root`'s child `name`, created after the last of `after` present (schema order) when missing. */
function child(root: XmlElement, name: XmlName, after: readonly XmlName[]): XmlElement {
  const found = root.element(name);
  if (found) return found;
  const created = new XmlElement(name);
  created.parent = root;
  const anchor = after.map((a) => root.element(a)).find((e) => e !== null);
  const at = anchor ? root.nodes.indexOf(anchor) + 1 : 0;
  root.nodes.splice(at, 0, created);
  return created;
}

const mm = (value: string | null): number | null => {
  const cm = parseLength(value);
  return cm === null ? null : cm * 10;
};

const SCALE_ATTRIBUTES = ['scale-to', 'scale-to-pages', 'scale-to-X', 'scale-to-Y'];
/** LibreOffice's extension namespace, where it writes the fit-to-width-and-height scale. */
const Loext = xmlNamespace('urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0', 'loext');

function readLayout(layout: XmlElement): PageSetup {
  const p = layout.element(Style('page-layout-properties'));
  const d = PageSetups.default;
  if (!p) return d;
  const margin = mm(p.attr(Fo('margin')));
  const side = (name: string, fallback: number) => mm(p.attr(Fo(name))) ?? margin ?? fallback;
  const print = (p.attr(Style('print')) ?? 'charts drawings objects zero-values').split(/\s+/);
  const centering = p.attr(Style('table-centering')) ?? 'none';
  const width = mm(p.attr(Fo('page-width'))) ?? d.paperWidth;
  const height = mm(p.attr(Fo('page-height'))) ?? d.paperHeight;
  const orientation = p.attr(Style('print-orientation')) ?? (width > height ? 'landscape' : 'portrait');
  return {
    paperWidth: round(width),
    paperHeight: round(height),
    orientation: orientation === 'landscape' ? 'landscape' : 'portrait',
    margins: {
      top: round(side('margin-top', d.margins.top)),
      bottom: round(side('margin-bottom', d.margins.bottom)),
      left: round(side('margin-left', d.margins.left)),
      right: round(side('margin-right', d.margins.right)),
    },
    scale: readScale(p),
    gridlines: print.includes('grid'),
    headers: print.includes('headers'),
    centerHorizontally: centering === 'horizontal' || centering === 'both',
    centerVertically: centering === 'vertical' || centering === 'both',
    pageOrder: p.attr(Style('print-page-order')) === 'ltr' ? 'overThenDown' : 'downThenOver',
  };
}

function readScale(p: XmlElement): PageScale {
  const int = (v: string | null) => (v === null ? null : Number.parseInt(v, 10));
  const x = int(p.attr(Style('scale-to-X')) ?? p.attr(Loext('scale-to-X')));
  const y = int(p.attr(Style('scale-to-Y')) ?? p.attr(Loext('scale-to-Y')));
  if ((x !== null && x > 0) || (y !== null && y > 0)) return { kind: 'fit', width: x ?? 0, height: y ?? 0 };
  const pages = int(p.attr(Style('scale-to-pages')));
  if (pages !== null && pages > 0) return { kind: 'pages', pages };
  const percent = Number.parseFloat(p.attr(Style('scale-to')) ?? '100');
  return { kind: 'percent', percent: Number.isFinite(percent) && percent > 0 ? percent : 100 };
}

function writeLayout(layout: XmlElement, setup: PageSetup): void {
  let p = layout.element(Style('page-layout-properties'));
  if (!p) {
    p = new XmlElement(Style('page-layout-properties'));
    layout.addFirst(p);
  }
  const length = (v: number) => `${formatUpTo3(v / 10)}cm`;
  p.setAttr(Fo('page-width'), length(setup.paperWidth))
    .setAttr(Fo('page-height'), length(setup.paperHeight))
    .setAttr(Style('print-orientation'), setup.orientation)
    .setAttr(Fo('margin'), null)
    .setAttr(Fo('margin-top'), length(setup.margins.top))
    .setAttr(Fo('margin-bottom'), length(setup.margins.bottom))
    .setAttr(Fo('margin-left'), length(setup.margins.left))
    .setAttr(Fo('margin-right'), length(setup.margins.right));
  // A setup's paper is no longer one of the printer's own named sizes.
  p.setAttr(Style('paper-tray-name'), null);
  for (const a of SCALE_ATTRIBUTES) p.setAttr(Style(a), null);
  p.setAttr(Loext('scale-to-X'), null).setAttr(Loext('scale-to-Y'), null);
  const s = setup.scale;
  if (s.kind === 'percent') p.setAttr(Style('scale-to'), `${formatUpTo3(s.percent)}%`);
  else if (s.kind === 'pages') p.setAttr(Style('scale-to-pages'), s.pages);
  else {
    // LibreOffice reads this pair from its own namespace; ODF 1.3 readers from style:.
    layout.declare(Loext.prefix, Loext.uri);
    p.setAttr(Style('scale-to-X'), s.width).setAttr(Style('scale-to-Y'), s.height);
    p.setAttr(Loext('scale-to-X'), s.width).setAttr(Loext('scale-to-Y'), s.height);
  }
  const tokens = (p.attr(Style('print')) ?? 'charts drawings objects zero-values')
    .split(/\s+/)
    .filter((t) => t && t !== 'grid' && t !== 'headers');
  if (setup.gridlines) tokens.push('grid');
  if (setup.headers) tokens.push('headers');
  p.setAttr(Style('print'), tokens.sort().join(' '));
  const h = setup.centerHorizontally;
  const v = setup.centerVertically;
  p.setAttr(Style('table-centering'), h && v ? 'both' : h ? 'horizontal' : v ? 'vertical' : 'none');
  p.setAttr(Style('print-page-order'), setup.pageOrder === 'overThenDown' ? 'ltr' : 'ttb');
}

/** To 0.1 mm: inches and centimetres round-trip without 14.99-style drift. */
const round = (v: number) => Math.round(v * 10) / 10;
