import type { CalculationMode } from '../../model/documentSettings.ts';
import type { Workbook } from '../../model/workbook.ts';
import { tryParseXml, XmlDocument, XmlElement, type XmlName } from '../../xml/index.ts';
import { Config, Dc, Fo, Manifest, OdsXmlNames as N, Office, Style } from './odsXmlNames.ts';

const CALC_MODE_ITEM = 'CellestraCalculationMode';
const LOCALE_ITEM = 'CellestraLocale';
const ACTIVE_TABLE_ITEM = 'ActiveTable';
const VIEWS_ITEM = 'Views';
const VIEW_ID_ITEM = 'ViewId';
const TABLES_ITEM = 'Tables';
const HORIZONTAL_SPLIT_MODE_ITEM = 'HorizontalSplitMode';
const VERTICAL_SPLIT_MODE_ITEM = 'VerticalSplitMode';
const HORIZONTAL_SPLIT_POSITION_ITEM = 'HorizontalSplitPosition';
const VERTICAL_SPLIT_POSITION_ITEM = 'VerticalSplitPosition';
const POSITION_RIGHT_ITEM = 'PositionRight';
const POSITION_BOTTOM_ITEM = 'PositionBottom';

/** ODF/LibreOffice split mode: 0 none, 1 a movable split, 2 a freeze. Only 2 is a freeze. */
const FROZEN_SPLIT_MODE = '2';
const NO_SPLIT_MODE = '0';

/** The calculation mode as the C# enum spelled it in settings.xml, so files stay interchangeable. */
const CALC_MODE_NAMES: Record<CalculationMode, string> = { automatic: 'Automatic', manual: 'Manual' };

/** Builds the smaller ODF parts: meta.xml, settings.xml, styles.xml, and the manifest. */
export const OdsSupportXml = {
  buildMeta(): XmlDocument {
    const meta = new XmlElement(
      Office('meta'),
      new XmlElement(Dc('date')),
      new XmlElement(Office('generator'), 'Cellestra/1.0'),
    );
    const root = new XmlElement(Office('document-meta'))
      .declare(Office.prefix, Office.uri)
      .declare(Dc.prefix, Dc.uri)
      .setAttr(Office('version'), '1.3')
      .add(meta);
    return new XmlDocument(root);
  },

  buildStyles(): XmlDocument {
    const root = new XmlElement(Office('document-styles'))
      .declare(Office.prefix, Office.uri)
      .declare(Style.prefix, Style.uri)
      .declare(Fo.prefix, Fo.uri)
      .setAttr(Office('version'), '1.3')
      .add(
        new XmlElement(Office('styles')),
        new XmlElement(Office('automatic-styles')),
        new XmlElement(Office('master-styles')),
      );
    return new XmlDocument(root);
  },

  /**
   * settings.xml for `workbook`. With the source file's settings (`previousXml`), what the model owns
   * — the active sheet, frozen panes, calculation mode, locale — is written into them, and the rest
   * (zoom, grid lines, the printer setup, every `ooo:configuration-settings` item) is kept as it was.
   */
  buildSettings(workbook: Workbook, previousXml: string | null = null): XmlDocument {
    const fresh = freshSettings(workbook);
    const root =
      previousXml === null ? null : (tryParseXml(previousXml, { preserveWhitespace: false })?.root ?? null);
    const settings = root?.is(Office('document-settings')) ? root.element(Office('settings')) : null;
    if (!root || !settings) return fresh;

    const freshSet = fresh.root.element(Office('settings'))!.element(Config('config-item-set'))!;
    const viewSet = named(settings, Config('config-item-set'), VIEW_SETTINGS);
    if (!viewSet) {
      settings.addFirst(freshSet);
      return new XmlDocument(root);
    }
    setItem(viewSet, CALC_MODE_ITEM, 'string', CALC_MODE_NAMES[workbook.settings.calculationMode]);
    setItem(viewSet, LOCALE_ITEM, 'string', workbook.settings.locale ?? '');

    const freshViews = freshSet.element(Config('config-item-map-indexed'))!;
    const views = named(viewSet, Config('config-item-map-indexed'), VIEWS_ITEM);
    const view = views?.element(Config('config-item-map-entry')) ?? null;
    if (!views) viewSet.add(freshViews);
    else if (!view) views.add(freshViews.element(Config('config-item-map-entry'))!);
    else {
      setItem(view, ACTIVE_TABLE_ITEM, 'string', activeSheetName(workbook));
      mergeTables(view, workbook);
    }
    return new XmlDocument(root);
  },

  applySettings(settingsXml: string, workbook: Workbook): void {
    const document = tryParseXml(settingsXml);
    if (!document) return;
    const root = document.root;

    for (const entry of root.descendants(Config('config-item-map-entry'))) {
      const sheetName = entry.attr(Config('name'));
      const sheet = sheetName === null ? undefined : workbook.findSheet(sheetName);
      if (!sheet) continue;
      sheet.frozenColumns = frozenLines(entry, HORIZONTAL_SPLIT_MODE_ITEM, HORIZONTAL_SPLIT_POSITION_ITEM);
      sheet.frozenRows = frozenLines(entry, VERTICAL_SPLIT_MODE_ITEM, VERTICAL_SPLIT_POSITION_ITEM);
    }

    for (const item of root.descendants(Config('config-item'))) {
      const name = item.attr(Config('name'));
      if (name === CALC_MODE_ITEM) {
        const mode = parseCalculationMode(item.value);
        if (mode !== null) workbook.settings.calculationMode = mode;
      } else if (name === LOCALE_ITEM) {
        // An explicit document locale round-trips; blank keeps "system default".
        workbook.settings.locale = item.value.trim() ? tryLocale(item.value) : null;
      } else if (name === ACTIVE_TABLE_ITEM) {
        const sheet = workbook.findSheet(item.value);
        if (sheet) workbook.settings.activeSheetIndex = workbook.indexOf(sheet);
      }
    }
  },

  buildManifest(
    parts: Iterable<string>,
    chartObjectDirs: Iterable<string> = [],
    previousManifest: string | null = null,
  ): XmlDocument {
    const manifest = new XmlElement(Manifest('manifest'))
      .declare(Manifest.prefix, Manifest.uri)
      .setAttr(Manifest('version'), '1.3');

    manifest.add(fileEntry('/', N.mimeType));

    // Each embedded chart sub-document needs its directory declared with the chart media type.
    const declared = new Set<string>();
    for (const dir of chartObjectDirs) {
      manifest.add(fileEntry(`${dir}/`, N.chartMediaType));
      declared.add(`${dir}/`);
    }

    // A preserved sub-document (an embedded formula, a drawing, a chart on a sheet the model does
    // not draw) is opened by its directory's media type. Regenerating the manifest from part names
    // alone dropped that entry, and LibreOffice showed the object as an empty frame.
    const partList = [...parts];
    for (const [dir, mediaType] of directoryEntries(previousManifest)) {
      if (declared.has(dir) || !partList.some((p) => p.startsWith(dir))) continue;
      manifest.add(fileEntry(dir, mediaType));
      declared.add(dir);
    }

    for (const part of partList) {
      if (part === N.mimetypePart) continue;
      manifest.add(fileEntry(part, mediaTypeFor(part)));
    }
    return new XmlDocument(manifest);
  },
} as const;

const VIEW_SETTINGS = 'ooo:view-settings';

function activeSheetName(workbook: Workbook): string {
  const count = workbook.sheetCount;
  return count > 0
    ? workbook.worksheets[Math.min(Math.max(workbook.settings.activeSheetIndex, 0), count - 1)]!.name
    : '';
}

/** settings.xml built from the model alone. */
function freshSettings(workbook: Workbook): XmlDocument {
  const activeSheet = activeSheetName(workbook);

  // Frozen panes and the active sheet live where LibreOffice looks for them: inside a view in
  // ooo:view-settings, with one Tables entry per sheet that has a freeze.
  const view = new XmlElement(
    Config('config-item-map-entry'),
    configItem(VIEW_ID_ITEM, 'string', 'view1'),
    tables(workbook),
    configItem(ACTIVE_TABLE_ITEM, 'string', activeSheet),
  );

  const itemSet = new XmlElement(Config('config-item-set'))
    .setAttr(Config('name'), VIEW_SETTINGS)
    .add(
      new XmlElement(Config('config-item-map-indexed')).setAttr(Config('name'), VIEWS_ITEM).add(view),
      configItem(CALC_MODE_ITEM, 'string', CALC_MODE_NAMES[workbook.settings.calculationMode]),
      configItem(LOCALE_ITEM, 'string', workbook.settings.locale ?? ''),
    );

  const root = new XmlElement(Office('document-settings'))
    .declare(Office.prefix, Office.uri)
    .declare(Config.prefix, Config.uri)
    .setAttr(Office('version'), '1.3')
    .add(new XmlElement(Office('settings'), itemSet));
  return new XmlDocument(root);
}

/** The child `name`d element (by `config:name`) of `parent`, or null. */
function named(parent: XmlElement, element: XmlName, name: string): XmlElement | null {
  return parent.elements(element).find((e) => e.attr(Config('name')) === name) ?? null;
}

/** Set (or add) the `config:config-item` called `name` under `parent`. */
function setItem(parent: XmlElement, name: string, type: string, value: string): void {
  const item = named(parent, Config('config-item'), name);
  if (item) item.value = value;
  else parent.add(configItem(name, type, value));
}

/**
 * Bring a view's per-sheet `Tables` map in line with the model: entries for sheets that no longer
 * exist go, and a sheet whose freeze changed gets the split items rewritten. An entry whose freeze
 * is unchanged keeps every item (scroll position, cursor, zoom) untouched.
 */
function mergeTables(view: XmlElement, workbook: Workbook): void {
  let map = named(view, Config('config-item-map-named'), TABLES_ITEM);
  if (!map)
    view.add((map = new XmlElement(Config('config-item-map-named')).setAttr(Config('name'), TABLES_ITEM)));
  for (const entry of map.elements(Config('config-item-map-entry'))) {
    const name = entry.attr(Config('name'));
    if (name === null || !workbook.findSheet(name)) map.remove(entry);
  }
  for (const sheet of workbook.worksheets) {
    const entry = named(map, Config('config-item-map-entry'), sheet.name);
    const columns = Math.max(0, sheet.frozenColumns);
    const rows = Math.max(0, sheet.frozenRows);
    if (!entry) {
      if (columns > 0 || rows > 0) map.add(tableEntry(sheet.name, columns, rows));
      continue;
    }
    const was = {
      columns: frozenLines(entry, HORIZONTAL_SPLIT_MODE_ITEM, HORIZONTAL_SPLIT_POSITION_ITEM),
      rows: frozenLines(entry, VERTICAL_SPLIT_MODE_ITEM, VERTICAL_SPLIT_POSITION_ITEM),
    };
    if (was.columns === columns && was.rows === rows) continue;
    for (const item of tableEntry(sheet.name, columns, rows).elements()) {
      setItem(entry, item.attr(Config('name'))!, item.attr(Config('type'))!, item.value);
    }
  }
}

/**
 * One `Tables` map holding the frozen-pane state of every sheet that has one. Sheets with no freeze
 * are left out — an absent entry already means "no panes".
 */
function tables(workbook: Workbook): XmlElement {
  const map = new XmlElement(Config('config-item-map-named')).setAttr(Config('name'), TABLES_ITEM);
  for (const sheet of workbook.worksheets) {
    if (sheet.frozenRows <= 0 && sheet.frozenColumns <= 0) continue;
    map.add(tableEntry(sheet.name, sheet.frozenColumns, sheet.frozenRows));
  }
  return map;
}

/** One sheet's `Tables` entry: its split modes and positions for a freeze of `columns` × `rows`. */
function tableEntry(name: string, columns: number, rows: number): XmlElement {
  return new XmlElement(Config('config-item-map-entry')).setAttr(Config('name'), name).add(
    configItem(HORIZONTAL_SPLIT_MODE_ITEM, 'short', columns > 0 ? FROZEN_SPLIT_MODE : NO_SPLIT_MODE),
    configItem(VERTICAL_SPLIT_MODE_ITEM, 'short', rows > 0 ? FROZEN_SPLIT_MODE : NO_SPLIT_MODE),
    configItem(HORIZONTAL_SPLIT_POSITION_ITEM, 'int', String(columns)),
    configItem(VERTICAL_SPLIT_POSITION_ITEM, 'int', String(rows)),
    // Where the unfrozen pane starts — the same line, since a freeze splits exactly there.
    configItem(POSITION_RIGHT_ITEM, 'int', String(columns)),
    configItem(POSITION_BOTTOM_ITEM, 'int', String(rows)),
  );
}

/** `Enum.TryParse<CalculationMode>`: case-sensitive name match in C#; accepted case-insensitively here. */
function parseCalculationMode(text: string): CalculationMode | null {
  const t = text.trim().toLowerCase();
  return t === 'automatic' || t === 'manual' ? t : null;
}

/** `CultureInfo.GetCultureInfo(name)` or null: a well-formed BCP 47 tag, canonicalized. */
function tryLocale(name: string): string | null {
  try {
    return Intl.getCanonicalLocales(name.trim())[0] ?? null;
  } catch {
    return null;
  }
}

/**
 * How many lines a pane holds frozen on one axis: the position, but only when the mode says the
 * split is a freeze (a mode-1 split is a draggable scroll divider, not frozen headers).
 */
function frozenLines(entry: XmlElement, modeItem: string, positionItem: string): number {
  if (itemValue(entry, modeItem) !== FROZEN_SPLIT_MODE) return 0;
  const text = itemValue(entry, positionItem)?.trim() ?? '';
  if (!/^[+-]?\d+$/.test(text)) return 0;
  const lines = Number(text);
  return lines > 0 && lines <= 2147483647 ? lines : 0;
}

function itemValue(entry: XmlElement, name: string): string | null {
  return (
    entry.elements(Config('config-item')).find((item) => item.attr(Config('name')) === name)?.value ?? null
  );
}

/** The directory entries (`Object 1/` and the like) of a manifest, with their media types. */
function directoryEntries(manifestXml: string | null): Array<[path: string, mediaType: string]> {
  if (manifestXml === null) return [];
  const root = tryParseXml(manifestXml, { preserveWhitespace: false })?.root;
  const out: Array<[string, string]> = [];
  for (const entry of root?.elements(Manifest('file-entry')) ?? []) {
    const path = entry.attr(Manifest('full-path'));
    const mediaType = entry.attr(Manifest('media-type'));
    if (path && path !== '/' && path.endsWith('/') && mediaType) out.push([path, mediaType]);
  }
  return out;
}

function fileEntry(path: string, mediaType: string): XmlElement {
  return new XmlElement(Manifest('file-entry'))
    .setAttr(Manifest('full-path'), path)
    .setAttr(Manifest('media-type'), mediaType);
}

function configItem(name: string, type: string, value: string): XmlElement {
  return new XmlElement(Config('config-item'), value)
    .setAttr(Config('name'), name)
    .setAttr(Config('type'), type);
}

function mediaTypeFor(part: string): string {
  const lower = part.toLowerCase();
  if (lower.endsWith('.xml')) return 'text/xml';
  if (lower.endsWith('.png')) return 'image/png';
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
  return 'application/octet-stream';
}
