import type { Workbook } from '../../model/workbook.ts';
import { serializeXml, tryParseXml } from '../../xml/index.ts';
import { writeContentXml } from './odsContentWrite.ts';
import { OdsCharts } from './odsCharts.ts';
import { OdsPackage } from './odsPackage.ts';
import { OdsSupportXml } from './odsSupportXml.ts';
import { OdsXmlNames as N, Style } from './odsXmlNames.ts';

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
    const reservedStyleNames = styleNames(pkg.getText(N.stylesPart));
    pkg.setText(
      N.contentPart,
      serializeXml(writeContentXml(workbook, pkg.foreignSpreadsheetChildren, reservedStyleNames), flat),
    );
    pkg.setText(
      N.settingsPart,
      serializeXml(OdsSupportXml.buildSettings(workbook, pkg.getText(N.settingsPart)), flat),
    );

    if (!pkg.contains(N.stylesPart))
      pkg.setText(N.stylesPart, serializeXml(OdsSupportXml.buildStyles(), flat));
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
function styleNames(stylesXml: string | null): Set<string> {
  const names = new Set<string>();
  const root =
    stylesXml === null ? null : (tryParseXml(stylesXml, { preserveWhitespace: false })?.root ?? null);
  if (!root) return names;
  for (const e of root.descendants()) {
    const name = e.attr(Style('name'));
    if (name !== null) names.add(name);
  }
  return names;
}
