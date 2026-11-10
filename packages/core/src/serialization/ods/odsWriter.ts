import type { Workbook } from '../../model/workbook.ts';
import { XmlElement, serializeXml, tryParseXml, type XmlDocument } from '../../xml/index.ts';
import { writeContentXml } from './odsContentWrite.ts';
import { OdsCharts } from './odsCharts.ts';
import { OdsPackage } from './odsPackage.ts';
import { OdsPageSetup, NO_PAGE_PLAN } from './odsPageSetup.ts';
import { OdsPreserved } from './odsPreserved.ts';
import { OdsSupportXml } from './odsSupportXml.ts';
import { CONDITIONAL_STYLE_PREFIX } from './odsStyleRegistry.ts';
import { Fo, OdsXmlNames as N, NumberNs, Office, Style } from './odsXmlNames.ts';

/**
 * Writes a `Workbook` to `.ods` bytes. Regenerates content.xml, settings.xml, and the manifest from
 * the model; preserves the source package's styles.xml, meta.xml, and all unmodeled parts and
 * `office:spreadsheet` children verbatim (PRD §8.1).
 */
export const OdsWriter = {
  write(workbook: Workbook, source: OdsPackage | null = null): Uint8Array {
    const pkg = source ?? new OdsPackage();

    pkg.setText(N.mimetypePart, N.mimeType);

    // Chart objects must exist (with assigned Object N names) before content.xml is built, since
    // the table writer embeds a draw frame pointing at each one.
    const chartObjectDirs = OdsCharts.writeObjects(workbook, pkg);

    // Never indented: whitespace between the spans of a `text:p` (a cell, a comment, a shape's
    // text) is text to a reader, and a reader cannot tell indentation from it.
    const flat = { indent: false } as const;
    const stylesText = pkg.getText(N.stylesPart);
    const styles =
      stylesText === null
        ? OdsSupportXml.buildStyles()
        : tryParseXml(stylesText, { preserveWhitespace: true });
    const stripped = styles ? dropConditionalStyles(styles) : false;
    // Page layouts first: the master pages they add are names the content writer must not reuse.
    const pages = styles
      ? OdsPageSetup.write(workbook, styles, OdsPreserved.workbook(workbook)?.automaticStyles ?? [])
      : NO_PAGE_PLAN;
    const reservedStyleNames = styles ? styleNames(styles) : new Set<string>();
    for (const s of pages.tableStyles) reservedStyleNames.add(s.attr(Style('name'))!);
    // Named styles conditional formats apply; null when styles.xml can't be read (left as it was).
    const commonStyles: XmlElement[] | null = styles ? [] : null;
    pkg.setText(
      N.contentPart,
      serializeXml(
        writeContentXml(workbook, pkg.foreignSpreadsheetChildren, reservedStyleNames, commonStyles, pages),
        flat,
      ),
    );
    if (styles && (stylesText === null || stripped || pages.stylesChanged || commonStyles!.length > 0)) {
      let officeStyles = styles.root.element(Office('styles'));
      if (!officeStyles) {
        // Schema order: office:font-face-decls, then office:styles.
        officeStyles = new XmlElement(Office('styles'));
        styles.root.addFirst(officeStyles);
        const fonts = styles.root.element(Office('font-face-decls'));
        if (fonts && styles.root.remove(fonts)) styles.root.addFirst(fonts);
      }
      if (commonStyles!.length > 0) {
        styles.root.declare(Style.prefix, Style.uri).declare(Fo.prefix, Fo.uri);
        if (commonStyles!.some((e) => e.uri === NumberNs.uri))
          styles.root.declare(NumberNs.prefix, NumberNs.uri);
        for (const e of commonStyles!) officeStyles.add(e);
      }
      pkg.setText(N.stylesPart, serializeXml(styles, flat));
    }
    pkg.setText(
      N.settingsPart,
      serializeXml(OdsSupportXml.buildSettings(workbook, pkg.getText(N.settingsPart)), flat),
    );

    if (!pkg.contains(N.metaPart)) pkg.setText(N.metaPart, serializeXml(OdsSupportXml.buildMeta(), flat));

    // The manifest lists every part (regenerated to match the final contents), plus a directory
    // entry per embedded chart object declared with the chart media type.
    const parts = pkg.partNames.filter((p) => p !== N.manifestPart);
    const previousManifest = pkg.getText(N.manifestPart);
    pkg.setText(
      N.manifestPart,
      serializeXml(OdsSupportXml.buildManifest(parts, chartObjectDirs, previousManifest), flat),
    );

    return pkg.write();
  },
} as const;

/**
 * Every `style:name` the preserved styles.xml declares. The automatic styles content.xml generates
 * (`ce1`, `N0`) must not reuse one: LibreOffice resolves a data style by name across both parts,
 * so a generated `N0` beside the page styles' own `N0` shows one of them with the other's format.
 */
function styleNames(styles: XmlDocument): Set<string> {
  const names = new Set<string>();
  for (const e of styles.root.descendants()) {
    const name = e.attr(Style('name'));
    if (name !== null) names.add(name);
  }
  return names;
}

/** Remove the named styles an earlier save wrote for conditional formats; true when there were any. */
function dropConditionalStyles(styles: XmlDocument): boolean {
  const officeStyles = styles.root.element(Office('styles'));
  if (!officeStyles) return false;
  let dropped = false;
  for (const e of officeStyles.elements(Style('style'))) {
    if (e.attr(Style('name'))?.startsWith(CONDITIONAL_STYLE_PREFIX))
      dropped = officeStyles.remove(e) || dropped;
  }
  return dropped;
}
