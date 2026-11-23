import { Defaults } from '../../defaults.ts';
import { OpenFormulaTranslator } from '../../formulas/openFormulaTranslator.ts';
import { CellRange } from '../../model/cellRange.ts';
import type { NamedRange } from '../../model/namedRange.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { XmlElement } from '../../xml/index.ts';
import { odfCellAddress, odfRangeAddress, odfSheetName } from './odsContentShared.ts';
import { Table } from './odsXmlNames.ts';

/**
 * Names in `content.xml`: `table:named-expressions`, global under `office:spreadsheet` and
 * sheet-local inside a `table:table`. A `table:named-range` holds cells and a
 * `table:named-expression` a formula (`Rate = of:=0.07`). What the model can't hold is kept
 * verbatim and written back unless the model now holds that name itself.
 */

/**
 * Read a `table:named-expressions` block: each named range or named expression the model can hold
 * becomes a name (sheet-local when `scope` is a sheet name); everything else — a range or a
 * formula the model cannot parse — is returned to be kept verbatim.
 */
export function readNamedExpressions(
  container: XmlElement,
  workbook: Workbook,
  scope: string | null,
): XmlElement[] {
  const preserved: XmlElement[] = [];
  const taken = (name: string): boolean => {
    const key = name.toLowerCase();
    return workbook.namedRanges.some(
      (n) => n.name.toLowerCase() === key && (n.scope ?? null)?.toLowerCase() === scope?.toLowerCase(),
    );
  };
  const add = (named: NamedRange): void => {
    if (!taken(named.name)) workbook.addNamedRange(scope === null ? named : { ...named, scope });
  };
  for (const named of container.elements()) {
    if (named.is(Table('named-expression'))) {
      const name = named.attr(Table('name'));
      const expression = named.attr(Table('expression'));
      if (name !== null && expression !== null) {
        const read = OpenFormulaTranslator.read(expression);
        if (read.source === null) {
          add({ name, expression: read.formula });
          continue;
        }
      }
    }
    if (named.is(Table('named-range'))) {
      const name = named.attr(Table('name'));
      const address = named.attr(Table('cell-range-address'));
      if (name !== null && address !== null) {
        const internalRef = OpenFormulaTranslator.fromOpenFormula(
          'of:=[' + address.replace(/^\$+/, '') + ']',
        );
        const range = CellRange.tryParse(internalRef);
        if (range) {
          add({ name, target: range });
          continue;
        }
      }
    }
    preserved.push(named.clone());
  }
  return preserved;
}

/**
 * The workbook's global names: every modeled name, then each preserved entry (one the model could
 * not read) whose name the model does not now hold itself.
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
    if (named.expression !== undefined) {
      container.add(
        new XmlElement(Table('named-expression'))
          .setAttr(Table('name'), named.name)
          .setAttr(Table('expression'), OpenFormulaTranslator.toOpenFormula(named.expression))
          .setAttr(Table('base-cell-address'), `$${odfSheetName(fallbackSheet)}.$A$1`),
      );
      continue;
    }
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
