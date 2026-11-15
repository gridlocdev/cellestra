import { Defaults } from '../../defaults.ts';
import { OpenFormulaTranslator } from '../../formulas/openFormulaTranslator.ts';
import { CellRange } from '../../model/cellRange.ts';
import type { NamedRange } from '../../model/namedRange.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { XmlElement } from '../../xml/index.ts';
import { odfCellAddress, odfRangeAddress } from './odsContentShared.ts';
import { Table } from './odsXmlNames.ts';

/**
 * Named ranges in `content.xml`: `table:named-expressions`, global under `office:spreadsheet` and
 * sheet-local inside a `table:table`. What the model can't hold (a named expression) is kept
 * verbatim and written back unless the model now holds that name itself.
 */

/**
 * Read a `table:named-expressions` block: each `table:named-range` the model can hold becomes a
 * named range (sheet-local when `scope` is a sheet name); everything else — a named expression
 * (`Rate = of:=0.07`), a range the model cannot parse — is returned to be kept verbatim.
 */
export function readNamedExpressions(
  container: XmlElement,
  workbook: Workbook,
  scope: string | null,
): XmlElement[] {
  const preserved: XmlElement[] = [];
  for (const named of container.elements()) {
    if (named.is(Table('named-range'))) {
      const name = named.attr(Table('name'));
      const address = named.attr(Table('cell-range-address'));
      if (name !== null && address !== null) {
        const internalRef = OpenFormulaTranslator.fromOpenFormula(
          'of:=[' + address.replace(/^\$+/, '') + ']',
        );
        const range = CellRange.tryParse(internalRef);
        if (range) {
          const key = name.toLowerCase();
          const taken = workbook.namedRanges.some(
            (n) => n.name.toLowerCase() === key && (n.scope ?? null)?.toLowerCase() === scope?.toLowerCase(),
          );
          if (!taken)
            workbook.addNamedRange(scope === null ? { name, target: range } : { name, target: range, scope });
          continue;
        }
      }
    }
    preserved.push(named.clone());
  }
  return preserved;
}

/**
 * The workbook's global names: every modeled named range, then each preserved entry (a named
 * expression such as `Rate = of:=0.07`) whose name the model does not now hold itself.
 */
export function addNamedExpressions(
  spreadsheet: XmlElement,
  workbook: Workbook,
  preserved: readonly XmlElement[],
): void {
  const global = workbook.namedRanges.filter((n) => n.scope === null || n.scope === undefined);
  const fallbackSheet =
    workbook.worksheets.length > 0 ? workbook.worksheets[0]!.name : Defaults.firstSheetName;
  const container = namedExpressionsElement(global, preserved, fallbackSheet);
  if (container) spreadsheet.add(container);
}

/** A sheet's own names (`scope` = the sheet), written inside its table as ODF scopes them. */
export function addLocalNamedExpressions(
  table: XmlElement,
  workbook: Workbook,
  sheet: Worksheet,
  preserved: readonly XmlElement[],
): void {
  const key = sheet.name.toLowerCase();
  const local = workbook.namedRanges.filter((n) => n.scope?.toLowerCase() === key);
  const container = namedExpressionsElement(local, preserved, sheet.name);
  if (container) table.add(container);
}

function namedExpressionsElement(
  modeled: readonly NamedRange[],
  preserved: readonly XmlElement[],
  fallbackSheet: string,
): XmlElement | null {
  const container = new XmlElement(Table('named-expressions'));
  const taken = new Set<string>();
  for (const named of modeled) {
    taken.add(named.name.toLowerCase());
    container.add(
      new XmlElement(Table('named-range'))
        .setAttr(Table('name'), named.name)
        .setAttr(Table('cell-range-address'), odfRangeAddress(named.target, fallbackSheet))
        .setAttr(Table('base-cell-address'), odfCellAddress(named.target.topLeft, fallbackSheet)),
    );
  }
  for (const element of preserved) {
    const name = element.attr(Table('name'));
    if (name !== null && taken.has(name.toLowerCase())) continue;
    container.add(element.clone());
  }
  return container.hasElements ? container : null;
}
