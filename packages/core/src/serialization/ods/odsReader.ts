import { Defaults } from '../../defaults.ts';
import { Workbook } from '../../model/workbook.ts';
import { parseXml, tryParseXml } from '../../xml/index.ts';
import { readContentXml } from './odsContentRead.ts';
import { OdsPackage } from './odsPackage.ts';
import { OdsPageSetup } from './odsPageSetup.ts';
import { OdsSupportXml } from './odsSupportXml.ts';
import { OdsXmlNames as N } from './odsXmlNames.ts';

/** A workbook read from a `.ods` file, with the package that backs it (for a lossless re-save). */
export interface OdsReadResult {
  readonly workbook: Workbook;
  readonly package: OdsPackage;
}

/**
 * Reads `.ods` bytes into a `Workbook`. The backing `OdsPackage` is returned too so a later save can
 * preserve unmodeled parts and content verbatim (PRD §8.1). Cached cell values load as-is (PRD §8.6).
 */
export const OdsReader = {
  readWithPackage(bytes: Uint8Array): OdsReadResult {
    const pkg = OdsPackage.read(bytes);
    return { workbook: OdsReader.fromPackage(pkg), package: pkg };
  },

  read(bytes: Uint8Array): Workbook {
    return OdsReader.readWithPackage(bytes).workbook;
  },

  fromPackage(pkg: OdsPackage): Workbook {
    const workbook = new Workbook();

    const contentXml = pkg.getText(N.contentPart);
    if (contentXml !== null) {
      const document = parseXml(contentXml, { preserveWhitespace: true });
      // styles.xml holds the common cell styles and, in LibreOffice files, many data styles.
      const stylesXml = pkg.getText(N.stylesPart);
      // Whitespace is kept: a data style's `<number:text> </number:text>` is the space in "d mmmm".
      const styles = stylesXml === null ? null : tryParseXml(stylesXml, { preserveWhitespace: true });
      readContentXml(document, workbook, pkg.foreignSpreadsheetChildren, pkg, styles);
      OdsPageSetup.read(workbook, document.root, styles?.root ?? null);
    }

    // A valid spreadsheet always has at least one sheet.
    if (workbook.sheetCount === 0) workbook.addSheet(Defaults.firstSheetName);

    const settingsXml = pkg.getText(N.settingsPart);
    if (settingsXml !== null) OdsSupportXml.applySettings(settingsXml, workbook);

    return workbook;
  },
} as const;
