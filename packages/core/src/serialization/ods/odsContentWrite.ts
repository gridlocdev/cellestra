import type { Workbook } from '../../model/workbook.ts';
import { parseXmlElement, XmlDocument, XmlElement } from '../../xml/index.ts';
import { Form, Ooo } from './odsCheckboxes.ts';
import { addDatabaseRanges } from './odsDatabaseRanges.ts';
import { addNamedExpressions } from './odsNamedExpressions.ts';
import type { ForeignXml } from './odsPackage.ts';
import { NO_PAGE_PLAN, isGeneratedTableStyle, type PagePlan } from './odsPageSetup.ts';
import { OdsPreserved } from './odsPreserved.ts';
import { StyleRegistry } from './odsStyleRegistry.ts';
import { writeTable, type ValidationTargets } from './odsTableWrite.ts';
import { addContentValidations } from './odsValidation.ts';
import {
  Calcext,
  Dc,
  Draw,
  Fo,
  NumberNs,
  Of,
  Office,
  Style,
  Svg,
  Table,
  Text,
  Xlink,
} from './odsXmlNames.ts';

const NO_VALIDATIONS: ValidationTargets = [];

/**
 * Build `content.xml`: the cell grid, automatic styles, merges, column widths/row heights, hidden
 * lines, named ranges, validations, and database ranges. Unmodeled `office:spreadsheet` children
 * captured on read are re-emitted verbatim in their original position relative to the tables, and
 * so is the table, cell and named-expression content kept by `OdsPreserved`.
 */
export function writeContentXml(
  workbook: Workbook,
  foreign: readonly ForeignXml[],
  reservedStyleNames: ReadonlySet<string> = new Set(),
  commonStyles: XmlElement[] | null = null,
  pages: PagePlan = NO_PAGE_PLAN,
): XmlDocument {
  const autoStyles = new XmlElement(Office('automatic-styles'));
  const preserved = OdsPreserved.workbook(workbook);
  // Generated names must not reuse a preserved automatic style's name, nor any styles.xml name.
  const reserved = new Set<string>(reservedStyleNames);
  for (const style of preserved?.automaticStyles ?? []) {
    // Table styles an earlier save wrote for page setups are written again below, renamed.
    if (isGeneratedTableStyle(style)) continue;
    const name = style.attr(Style('name'));
    if (name !== null) reserved.add(name);
    autoStyles.add(style.clone());
  }
  for (const style of pages.tableStyles) autoStyles.add(style.clone());
  const registry = new StyleRegistry(workbook.styles, autoStyles, reserved, reservedStyleNames, commonStyles);

  const spreadsheet = new XmlElement(Office('spreadsheet'));
  // Structure protection and its key live here; the model has no field for them.
  for (const a of preserved?.spreadsheetAttributes ?? []) spreadsheet.setAttr(a.name, a.value);

  for (const child of foreign) if (child.beforeTables) spreadsheet.add(parseXmlElement(child.xml));

  // Data validations are declared once, before the tables (ODF schema order), and referenced by
  // name from each validated cell.
  const validationNames = addContentValidations(spreadsheet, workbook);

  workbook.worksheets.forEach((sheet, index) => {
    spreadsheet.add(
      writeTable(workbook, sheet, index, registry, validationNames.get(sheet.name) ?? NO_VALIDATIONS, pages),
    );
  });

  addNamedExpressions(spreadsheet, workbook, preserved?.namedExpressions ?? []);
  addDatabaseRanges(spreadsheet, workbook);

  for (const child of foreign) if (!child.beforeTables) spreadsheet.add(parseXmlElement(child.xml));

  const root = new XmlElement(Office('document-content'))
    .declare(Office.prefix, Office.uri)
    .declare(Table.prefix, Table.uri)
    .declare(Text.prefix, Text.uri)
    .declare(Style.prefix, Style.uri)
    .declare(Fo.prefix, Fo.uri)
    .declare(NumberNs.prefix, NumberNs.uri)
    .declare(Svg.prefix, Svg.uri)
    .declare(Draw.prefix, Draw.uri)
    .declare(Form.prefix, Form.uri)
    .declare(Ooo.prefix, Ooo.uri)
    .declare(Xlink.prefix, Xlink.uri)
    .declare(Of.prefix, Of.uri)
    .declare(Calcext.prefix, Calcext.uri)
    // Comments (office:annotation) carry dc:creator and dc:date; declared once here, not per note.
    .declare(Dc.prefix, Dc.uri)
    .setAttr(Office('version'), '1.3')
    // Schema order: office:scripts, office:font-face-decls, office:automatic-styles, office:body.
    .add(
      preserved?.scripts?.clone(),
      preserved?.fontFaces?.clone(),
      autoStyles,
      new XmlElement(Office('body'), spreadsheet),
    );

  return new XmlDocument(root);
}
