import type { Workbook } from '../../model/workbook.ts';
import { serializeXml } from '../../xml/index.ts';
import { writeContentXml } from './odsContentWrite.ts';
import { OdsCharts } from './odsCharts.ts';
import { OdsPackage } from './odsPackage.ts';
import { OdsSupportXml } from './odsSupportXml.ts';
import { OdsXmlNames as N } from './odsXmlNames.ts';

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

    pkg.setText(N.contentPart, serializeXml(writeContentXml(workbook, pkg.foreignSpreadsheetChildren)));
    pkg.setText(N.settingsPart, serializeXml(OdsSupportXml.buildSettings(workbook)));

    if (!pkg.contains(N.stylesPart)) pkg.setText(N.stylesPart, serializeXml(OdsSupportXml.buildStyles()));
    if (!pkg.contains(N.metaPart)) pkg.setText(N.metaPart, serializeXml(OdsSupportXml.buildMeta()));

    // The manifest lists every part (regenerated to match the final contents), plus a directory
    // entry per embedded chart object declared with the chart media type.
    const parts = pkg.partNames.filter((p) => p !== N.manifestPart);
    pkg.setText(N.manifestPart, serializeXml(OdsSupportXml.buildManifest(parts, chartObjectDirs)));

    return pkg.write();
  },
} as const;
