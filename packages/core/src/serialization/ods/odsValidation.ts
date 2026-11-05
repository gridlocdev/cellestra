import {
  dataValidation,
  type DataValidation,
  type ValidationKind,
  type ValidationOperator,
  type ValidationResponse,
} from '../../data/types.ts';
import { OpenFormulaTranslator } from '../../formulas/openFormulaTranslator.ts';
import { CellAddress } from '../../model/cellAddress.ts';
import { CellRange } from '../../model/cellRange.ts';
import { cellKey } from '../../model/sheetDimensions.ts';
import type { Workbook } from '../../model/workbook.ts';
import { roundTripNumber } from '../../values/cellValue.ts';
import { XmlElement } from '../../xml/index.ts';
import { parseInvariantFloat } from './mapping/odsCellStyles.ts';
import { odfCellAddress, odfRangeAddress, type ValidationNames } from './odsContentShared.ts';
import { OdsXmlNames as N, Table, Text } from './odsXmlNames.ts';

// ---- table:content-validation (data validation, PRD §8.13) -----------------

/**
 * Emit a `table:content-validations` block (one `table:content-validation` per rule across all
 * sheets) and return, per sheet, a map from each validated cell to the rule's name. The rule's
 * condition encodes its kind (whole-number/decimal/date/text-length/list), operator, operands and
 * list values in OpenFormula syntax; the error response maps to the `table:error-message`'s
 * `message-type`. Writes no block when nothing is validated (PRD §8.13).
 */
export function addContentValidations(spreadsheet: XmlElement, workbook: Workbook): ValidationNames {
  const perSheet = new Map<string, Map<number, string>>();
  const container = new XmlElement(Table(N.contentValidations));
  let next = 0;

  for (const sheet of workbook.worksheets) {
    if (sheet.validations.length === 0) continue;
    const cellMap = new Map<number, string>();
    for (const validation of sheet.validations) {
      const name = `val${next++}`;
      container.add(writeContentValidation(name, validation, sheet.name));
      // Last rule wins on overlap (rules should not overlap).
      for (const [row, column] of validation.target.positions()) cellMap.set(cellKey(row, column), name);
    }
    perSheet.set(sheet.name, cellMap);
  }

  if (container.hasElements) spreadsheet.add(container);
  return perSheet;
}

function writeContentValidation(name: string, validation: DataValidation, sheetName: string): XmlElement {
  const element = new XmlElement(Table(N.contentValidation))
    .setAttr(Table('name'), name)
    .setAttr(Table(N.condition), buildCondition(validation, sheetName))
    .setAttr(Table(N.allowEmptyCell), 'true')
    .setAttr(Table(N.baseCellAddress), odfCellAddress(validation.target.topLeft, sheetName));

  if (validation.kind === 'list') element.setAttr(Table(N.displayList), 'unsorted');

  const input = validation.inputMessage;
  if (input) {
    element.add(
      new XmlElement(Table(N.helpMessage))
        .setAttr(Table(N.display), 'true')
        .add(new XmlElement(Text('p'), input)),
    );
  }

  // The error response: Stop/Warning/Information map to the ODF message-type; Stop rejects.
  const error = new XmlElement(Table(N.errorMessage))
    .setAttr(Table(N.display), 'true')
    .setAttr(Table(N.messageType), messageTypeOf(validation.response));
  if (validation.errorMessage) error.add(new XmlElement(Text('p'), validation.errorMessage));
  element.add(error);

  return element;
}

function messageTypeOf(response: ValidationResponse): string {
  return response === 'warning' ? 'warning' : response === 'information' ? 'information' : 'stop';
}

function responseOf(messageType: string | null): ValidationResponse {
  return messageType === 'warning' ? 'warning' : messageType === 'information' ? 'information' : 'stop';
}

function buildCondition(validation: DataValidation, sheetName: string): string {
  if (validation.kind === 'list') {
    const source = validation.listSource;
    if (source) {
      return `of:cell-content-is-in-list([${odfRangeAddress(source.withSheet(source.sheet ?? sheetName), sheetName)}])`;
    }
    const items = (validation.listValues ?? []).map((v) => '"' + v.replaceAll('"', '""') + '"').join(';');
    return `of:cell-content-is-in-list(${items})`;
  }

  const kindPredicate =
    validation.kind === 'wholeNumber'
      ? 'of:cell-content-is-whole-number() and '
      : validation.kind === 'decimal'
        ? 'of:cell-content-is-decimal-number() and '
        : validation.kind === 'date'
          ? 'of:cell-content-is-date() and '
          : ''; // TextLength carries its kind in the content() call below

  const content = validation.kind === 'textLength' ? 'of:cell-content-text-length()' : 'of:cell-content()';
  const op1 = roundTripNumber(validation.operand1);
  const op2 = roundTripNumber(validation.operand2);

  let comparison: string;
  switch (validation.operator) {
    case 'between':
      comparison = `of:cell-content-is-between(${op1},${op2})`;
      break;
    case 'notBetween':
      comparison = `of:cell-content-is-not-between(${op1},${op2})`;
      break;
    case 'equal':
      comparison = `${content}=${op1}`;
      break;
    case 'notEqual':
      comparison = `${content}!=${op1}`;
      break;
    case 'greaterThan':
      comparison = `${content}>${op1}`;
      break;
    case 'lessThan':
      comparison = `${content}<${op1}`;
      break;
    case 'greaterOrEqual':
      comparison = `${content}>=${op1}`;
      break;
    case 'lessOrEqual':
      comparison = `${content}<=${op1}`;
      break;
    default:
      comparison = `${content}=${op1}`;
  }

  // Text-length between uses the text-length variant of the between predicate.
  if (validation.kind === 'textLength') {
    if (validation.operator === 'between')
      comparison = `of:cell-content-text-length-is-between(${op1},${op2})`;
    else if (validation.operator === 'notBetween')
      comparison = `of:cell-content-text-length-is-not-between(${op1},${op2})`;
  }

  return kindPredicate + comparison;
}

// ---- table:content-validation (read side) ----------------------------------

/** A parsed validation rule sans its target range (the target comes from the cells). */
export interface ValidationRule {
  readonly kind: ValidationKind;
  readonly operator: ValidationOperator;
  readonly operand1: number;
  readonly operand2: number;
  readonly listValues: readonly string[] | null;
  readonly listSource: CellRange | null;
  readonly inputMessage: string | null;
  readonly errorMessage: string | null;
  readonly response: ValidationResponse;
}

/** Per (sheet, rule name): the cells that referenced the rule. Keyed `sheet\u0000name`. */
export type ValidationCells = Map<
  string,
  { sheet: string; name: string; positions: Array<[number, number]> }
>;

export function readContentValidationRules(spreadsheet: XmlElement): Map<string, ValidationRule> {
  const rules = new Map<string, ValidationRule>();
  const container = spreadsheet.element(Table(N.contentValidations));
  if (!container) return rules;

  for (const element of container.elements(Table(N.contentValidation))) {
    const name = element.attr(Table('name'));
    if (name === null) continue;
    const condition = element.attr(Table(N.condition)) ?? '';
    const errorElement = element.element(Table(N.errorMessage));
    const helpElement = element.element(Table(N.helpMessage));
    const parsed = parseCondition(condition);
    rules.set(name, {
      ...parsed,
      inputMessage: paragraphText(helpElement),
      errorMessage: paragraphText(errorElement),
      response: responseOf(errorElement?.attr(Table(N.messageType)) ?? null),
    });
  }
  return rules;
}

function paragraphText(messageElement: XmlElement | null): string | null {
  if (!messageElement) return null;
  const joined = messageElement
    .elements(Text('p'))
    .map((p) => p.value)
    .join('\n');
  return joined.length > 0 ? joined : null;
}

type ParsedCondition = Pick<
  ValidationRule,
  'kind' | 'operator' | 'operand1' | 'operand2' | 'listValues' | 'listSource'
>;

function parseCondition(raw: string): ParsedCondition {
  const condition = raw.trim();
  const inList = 'of:cell-content-is-in-list(';
  if (condition.startsWith(inList) && condition.endsWith(')')) {
    const inner = condition.slice(inList.length, -1);
    if (inner.startsWith('[') && inner.endsWith(']')) {
      const internalRef = OpenFormulaTranslator.fromOpenFormula('of:=' + inner);
      const source = CellRange.tryParse(internalRef);
      return {
        kind: 'list',
        operator: 'between',
        operand1: 0,
        operand2: 0,
        listValues: null,
        listSource: source,
      };
    }
    return {
      kind: 'list',
      operator: 'between',
      operand1: 0,
      operand2: 0,
      listValues: parseQuotedList(inner),
      listSource: null,
    };
  }

  // Otherwise a number/date/text-length constraint. The leading predicate (if any) gives the kind;
  // the comparison gives the operator and operands.
  let kind: ValidationKind = 'decimal';
  if (condition.includes('cell-content-is-whole-number')) kind = 'wholeNumber';
  else if (condition.includes('cell-content-is-date')) kind = 'date';
  else if (condition.includes('cell-content-text-length')) kind = 'textLength';

  // The comparison clause is whatever follows the last " and " (or the whole string).
  const andIndex = condition.lastIndexOf(' and ');
  const comparison = andIndex >= 0 ? condition.slice(andIndex + 5) : condition;
  const [operator, operand1, operand2] = parseComparison(comparison);
  return { kind, operator, operand1, operand2, listValues: null, listSource: null };
}

function parseComparison(raw: string): [ValidationOperator, number, number] {
  const comparison = raw.trim();
  for (const [token, op] of [
    ['is-not-between(', 'notBetween'],
    ['is-between(', 'between'],
  ] as const) {
    const start = comparison.indexOf(token);
    if (start >= 0) {
      const open = start + token.length;
      const close = comparison.indexOf(')', open);
      if (close > open) {
        const parts = comparison.slice(open, close).split(',');
        const a = parts.length > 0 ? parseInvariant(parts[0]!) : 0;
        const b = parts.length > 1 ? parseInvariant(parts[1]!) : 0;
        return [op, a, b];
      }
    }
  }

  // Relational: cell-content()OP value (or cell-content-text-length()OP value).
  const parens = comparison.indexOf('()');
  const tail = parens >= 0 ? comparison.slice(parens + 2).trim() : comparison;
  for (const [prefix, op] of [
    ['>=', 'greaterOrEqual'],
    ['<=', 'lessOrEqual'],
    ['!=', 'notEqual'],
    ['<>', 'notEqual'],
    ['>', 'greaterThan'],
    ['<', 'lessThan'],
    ['=', 'equal'],
  ] as const) {
    if (tail.startsWith(prefix)) return [op, parseInvariant(tail.slice(prefix.length)), 0];
  }
  return ['between', 0, 0];
}

function parseQuotedList(inner: string): string[] {
  const values: string[] = [];
  let i = 0;
  while (i < inner.length) {
    // Skip separators/whitespace between items.
    while (i < inner.length && (inner[i] === ';' || /\s/.test(inner[i]!))) i++;
    if (i >= inner.length) break;

    if (inner[i] === '"') {
      i++;
      let sb = '';
      while (i < inner.length) {
        if (inner[i] === '"') {
          if (i + 1 < inner.length && inner[i + 1] === '"') {
            sb += '"';
            i += 2;
            continue;
          }
          i++;
          break;
        }
        sb += inner[i++];
      }
      values.push(sb);
    } else {
      const start = i;
      while (i < inner.length && inner[i] !== ';') i++;
      values.push(inner.slice(start, i).trim());
    }
  }
  return values;
}

/** `double.TryParse(…, NumberStyles.Any, Invariant)` or 0: thousands commas tolerated. */
function parseInvariant(text: string): number {
  return parseInvariantFloat(text.trim().replaceAll(',', '')) ?? 0;
}

export function recordValidationCell(
  cells: ValidationCells,
  sheet: string,
  name: string,
  row: number,
  column: number,
): void {
  const key = `${sheet}\u0000${name}`;
  let entry = cells.get(key);
  if (!entry) {
    entry = { sheet, name, positions: [] };
    cells.set(key, entry);
  }
  entry.positions.push([row, column]);
}

export function assembleValidations(
  workbook: Workbook,
  rules: ReadonlyMap<string, ValidationRule>,
  cells: ValidationCells,
): void {
  for (const { sheet: sheetName, name, positions } of cells.values()) {
    const rule = rules.get(name);
    if (!rule || positions.length === 0) continue;
    const sheet = workbook.findSheet(sheetName);
    if (!sheet) continue;

    // The bounding rectangle of all cells that referenced the rule is its target range.
    let top = Infinity;
    let bottom = -Infinity;
    let left = Infinity;
    let right = -Infinity;
    for (const [r, c] of positions) {
      top = Math.min(top, r);
      bottom = Math.max(bottom, r);
      left = Math.min(left, c);
      right = Math.max(right, c);
    }
    const target = new CellRange(new CellAddress(left, top), new CellAddress(right, bottom));

    sheet.addValidation(
      dataValidation(rule.kind, target, {
        operator: rule.operator,
        operand1: rule.operand1,
        operand2: rule.operand2,
        listValues: rule.listValues,
        listSource: rule.listSource,
        inputMessage: rule.inputMessage,
        errorMessage: rule.errorMessage,
        response: rule.response,
      }),
    );
  }
}
