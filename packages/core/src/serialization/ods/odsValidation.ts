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
import type { Workbook } from '../../model/workbook.ts';
import { roundTripNumber } from '../../values/cellValue.ts';
import { XmlElement } from '../../xml/index.ts';
import {
  appendParagraphText,
  odfCellAddress,
  odfRangeAddress,
  paragraphText,
  type ValidationNames,
} from './odsContentShared.ts';
import { OdsPreserved } from './odsPreserved.ts';
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
  const perSheet = new Map<string, Array<{ range: CellRange; name: string }>>();
  const container = new XmlElement(Table(N.contentValidations));
  let next = 0;

  for (const sheet of workbook.worksheets) {
    if (sheet.validations.length === 0) continue;
    const targets: Array<{ range: CellRange; name: string }> = [];
    for (const validation of sheet.validations) {
      const name = `val${next++}`;
      container.add(writeContentValidation(name, validation, sheet.name));
      // Later rules win on overlap (rules should not overlap): the writer looks targets up last-first.
      targets.push({ range: validation.target, name });
    }
    perSheet.set(sheet.name, targets);
  }

  if (container.hasElements) spreadsheet.add(container);
  return perSheet;
}

function writeContentValidation(name: string, validation: DataValidation, sheetName: string): XmlElement {
  // A rule read from a file whose condition the model cannot express goes back exactly as it came,
  // under its new name — its own base cell included, since relative references in the condition
  // are relative to it.
  const preserved = OdsPreserved.validation(validation);
  if (preserved) return preserved.element.clone().setAttr(Table('name'), name);

  // What the model has no field for comes from the element the rule was read from, if any.
  const source = OdsPreserved.validationSource(validation);
  const sourceHelp = source?.element(Table(N.helpMessage)) ?? null;
  const sourceError = source?.element(Table(N.errorMessage)) ?? null;

  const element = new XmlElement(Table(N.contentValidation)).setAttr(Table('name'), name);
  // "Any value" is the absence of a condition.
  if (validation.kind !== 'any') element.setAttr(Table(N.condition), buildCondition(validation, sheetName));
  element
    .setAttr(Table(N.allowEmptyCell), 'true')
    .setAttr(Table(N.baseCellAddress), odfCellAddress(validation.target.topLeft, sheetName));

  if (validation.kind === 'list') element.setAttr(Table(N.displayList), 'unsorted');
  copyAttributes(source ?? null, element, [
    N.condition,
    N.baseCellAddress,
    ...(validation.kind === 'list' ? [] : [N.displayList]),
  ]);

  const input = validation.inputMessage;
  if (input || sourceHelp) {
    const help = new XmlElement(Table(N.helpMessage)).setAttr(Table(N.display), 'true');
    copyAttributes(sourceHelp, help, []);
    addMessage(help, input);
    element.add(help);
  }

  // The error response: Stop/Warning/Information map to the ODF message-type; Stop rejects.
  const error = new XmlElement(Table(N.errorMessage)).setAttr(Table(N.display), 'true');
  copyAttributes(sourceError, error, [N.messageType]);
  error.setAttr(Table(N.messageType), messageTypeOf(validation.response));
  addMessage(error, validation.errorMessage);
  element.add(error);

  return element;
}

/** Copy `from`'s attributes onto `to` (overwriting), except `table:name` and the `skip`ped ones. */
function copyAttributes(from: XmlElement | null, to: XmlElement, skip: readonly string[]): void {
  for (const a of from?.attributes ?? []) {
    if (a.name.uri === Table.uri && (a.name.local === 'name' || skip.includes(a.name.local))) continue;
    to.setAttr(a.name, a.value);
  }
}

/** A message as one `text:p` per line (ODF has no newline inside a paragraph). */
function addMessage(target: XmlElement, message: string | null | undefined): void {
  if (!message) return;
  for (const line of message.replace(/\r\n?/g, '\n').split('\n')) {
    const paragraph = new XmlElement(Text('p'));
    appendParagraphText(paragraph, line);
    target.add(paragraph);
  }
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
      ? 'cell-content-is-whole-number() and '
      : validation.kind === 'decimal'
        ? 'cell-content-is-decimal-number() and '
        : validation.kind === 'date'
          ? 'cell-content-is-date() and '
          : ''; // TextLength carries its kind in the content() call below

  const content = validation.kind === 'textLength' ? 'cell-content-text-length()' : 'cell-content()';
  const op1 = roundTripNumber(validation.operand1);
  const op2 = roundTripNumber(validation.operand2);

  let comparison: string;
  switch (validation.operator) {
    case 'between':
      comparison = `cell-content-is-between(${op1},${op2})`;
      break;
    case 'notBetween':
      comparison = `cell-content-is-not-between(${op1},${op2})`;
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
    if (validation.operator === 'between') comparison = `cell-content-text-length-is-between(${op1},${op2})`;
    else if (validation.operator === 'notBetween')
      comparison = `cell-content-text-length-is-not-between(${op1},${op2})`;
  }

  // The namespace prefix opens the whole condition, once: LibreOffice reads
  // "of:cell-content-is-whole-number() and of:cell-content-is-between(1,9)" as no condition at all.
  return 'of:' + kindPredicate + comparison;
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
  /** The original element, when the condition is one the model cannot express. */
  readonly preserved: XmlElement | null;
  /** The original element, for what the model has no field for (titles, display flags). */
  readonly source: XmlElement;
}

/** A rectangle of cells, inclusive. */
interface Rect {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

/**
 * Per (sheet, rule name): the rectangles of cells that named the rule, in reading order (row
 * runs top to bottom, cells left to right). Keyed `sheet\u0000name`.
 */
export type ValidationCells = Map<string, { sheet: string; name: string; rects: Rect[] }>;

export function readContentValidationRules(spreadsheet: XmlElement): Map<string, ValidationRule> {
  const rules = new Map<string, ValidationRule>();
  const container = spreadsheet.element(Table(N.contentValidations));
  if (!container) return rules;

  for (const element of container.elements(Table(N.contentValidation))) {
    const name = element.attr(Table('name'));
    if (name === null) continue;
    const condition = element.attr(Table(N.condition));
    const errorElement = element.element(Table(N.errorMessage));
    const helpElement = element.element(Table(N.helpMessage));
    const parsed = condition === null ? ANY : parseCondition(condition);
    rules.set(name, {
      ...(parsed ?? ANY),
      inputMessage: messageText(helpElement),
      errorMessage: messageText(errorElement),
      response: responseOf(errorElement?.attr(Table(N.messageType)) ?? null),
      // Anything unrecognised is kept whole, rather than guessed at: a guess (the old "decimal
      // between 0 and 0") rejects every entry the original rule allowed.
      preserved: parsed === null ? element.clone() : null,
      source: element.clone(),
    });
  }
  return rules;
}

/** A message's text: its paragraphs' text (spaces, tabs and line breaks kept), one line each. */
function messageText(messageElement: XmlElement | null): string | null {
  if (!messageElement) return null;
  const joined = messageElement.elements(Text('p')).map(paragraphText).join('\n');
  return joined.length > 0 ? joined : null;
}

type ParsedCondition = Pick<
  ValidationRule,
  'kind' | 'operator' | 'operand1' | 'operand2' | 'listValues' | 'listSource'
>;

const ANY: ParsedCondition = {
  kind: 'any',
  operator: 'between',
  operand1: 0,
  operand2: 0,
  listValues: null,
  listSource: null,
};

const NUMBER = String.raw`[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?`;
const PREDICATE = /^(?:of:)?cell-content-is-(whole-number|decimal-number|date)\(\)\s+and\s+(.+)$/s;
const RELATIONAL = new RegExp(
  String.raw`^(?:of:)?cell-content(-text-length)?\(\)\s*(>=|<=|!=|<>|>|<|=)\s*(${NUMBER})$`,
);
const BETWEEN = new RegExp(
  String.raw`^(?:of:)?cell-content(-text-length)?-is-(not-)?between\(\s*(${NUMBER})\s*,\s*(${NUMBER})\s*\)$`,
);

const RELATIONAL_OPERATORS: Record<string, ValidationOperator> = {
  '>=': 'greaterOrEqual',
  '<=': 'lessOrEqual',
  '!=': 'notEqual',
  '<>': 'notEqual',
  '>': 'greaterThan',
  '<': 'lessThan',
  '=': 'equal',
};

/**
 * Parse the conditions the model can hold exactly; null for anything else (a formula condition,
 * a time rule, operands that are cell references or expressions, a list from a named range).
 */
function parseCondition(raw: string): ParsedCondition | null {
  const condition = raw.trim();
  const inList = /^(?:of:)?cell-content-is-in-list\((.*)\)$/s.exec(condition);
  if (inList) {
    const inner = inList[1]!.trim();
    if (inner.startsWith('[') && inner.endsWith(']') && !inner.slice(1, -1).includes('[')) {
      const internalRef = OpenFormulaTranslator.fromOpenFormula('of:=' + inner);
      const source = CellRange.tryParse(internalRef);
      return source ? { ...ANY, kind: 'list', listSource: source } : null;
    }
    const values = parseQuotedList(inner);
    return values ? { ...ANY, kind: 'list', listValues: values } : null;
  }

  // A number/date/text-length constraint: an optional kind predicate, then one comparison.
  let kind: ValidationKind | null = null;
  let comparison = condition;
  const predicate = PREDICATE.exec(condition);
  if (predicate) {
    kind = predicate[1] === 'whole-number' ? 'wholeNumber' : predicate[1] === 'date' ? 'date' : 'decimal';
    comparison = predicate[2]!.trim();
  }

  const between = BETWEEN.exec(comparison);
  const relational = between ? null : RELATIONAL.exec(comparison);
  const match = between ?? relational;
  if (!match) return null;
  const textLength = match[1] !== undefined;
  // A kind predicate on a text-length comparison is a combination the model has no word for.
  if (textLength && kind !== null) return null;
  const resolvedKind: ValidationKind = textLength ? 'textLength' : (kind ?? 'decimal');

  if (between) {
    return {
      ...ANY,
      kind: resolvedKind,
      operator: between[2] ? 'notBetween' : 'between',
      operand1: Number(between[3]),
      operand2: Number(between[4]),
    };
  }
  return {
    ...ANY,
    kind: resolvedKind,
    operator: RELATIONAL_OPERATORS[relational![2]!]!,
    operand1: Number(relational![3]),
  };
}

/** `"a";"b";3` → its values; null when an item is neither a quoted string nor a number. */
function parseQuotedList(inner: string): string[] | null {
  const values: string[] = [];
  let i = 0;
  while (i < inner.length) {
    // Skip separators/whitespace between items.
    while (i < inner.length && (inner[i] === ';' || /\s/.test(inner[i]!))) i++;
    if (i >= inner.length) break;

    if (inner[i] === '"') {
      i++;
      let sb = '';
      let closed = false;
      while (i < inner.length) {
        if (inner[i] === '"') {
          if (i + 1 < inner.length && inner[i + 1] === '"') {
            sb += '"';
            i += 2;
            continue;
          }
          i++;
          closed = true;
          break;
        }
        sb += inner[i++];
      }
      if (!closed) return null;
      values.push(sb);
    } else {
      const start = i;
      while (i < inner.length && inner[i] !== ';') i++;
      const item = inner.slice(start, i).trim();
      if (!new RegExp(`^${NUMBER}$`).test(item)) return null;
      values.push(item);
    }
  }
  return values;
}

/** Record that the cells in rows top..bottom, columns left..right name rule `name`. */
export function recordValidationRect(
  cells: ValidationCells,
  sheet: string,
  name: string,
  top: number,
  bottom: number,
  left: number,
  right: number,
): void {
  const key = `${sheet}\u0000${name}`;
  let entry = cells.get(key);
  if (!entry) {
    entry = { sheet, name, rects: [] };
    cells.set(key, entry);
  }
  // Neighbouring cells of one row run arrive one after another: widen the last rectangle.
  const last = entry.rects.at(-1);
  if (last && last.top === top && last.bottom === bottom && last.right + 1 === left) last.right = right;
  else entry.rects.push({ top, bottom, left, right });
}

/**
 * Stack the row-run rectangles into the fewest exact rectangles: a run directly below one with the
 * same columns extends it. The result covers exactly the cells that named the rule, never the
 * blank cells between them.
 */
function coalesce(rects: readonly Rect[]): Rect[] {
  const out: Rect[] = [];
  const open = new Map<string, Rect>();
  for (const r of rects) {
    const key = `${r.left}:${r.right}`;
    const above = open.get(key);
    if (above && above.bottom + 1 === r.top) {
      above.bottom = r.bottom;
      continue;
    }
    const copy = { ...r };
    out.push(copy);
    open.set(key, copy);
  }
  return out;
}

export function assembleValidations(
  workbook: Workbook,
  rules: ReadonlyMap<string, ValidationRule>,
  cells: ValidationCells,
): void {
  for (const { sheet: sheetName, name, rects } of cells.values()) {
    const rule = rules.get(name);
    if (!rule || rects.length === 0) continue;
    const sheet = workbook.findSheet(sheetName);
    if (!sheet) continue;

    // The model holds one rectangle per rule, so a rule over several becomes one rule each.
    for (const r of coalesce(rects)) {
      const target = new CellRange(new CellAddress(r.left, r.top), new CellAddress(r.right, r.bottom));
      const validation = dataValidation(rule.kind, target, {
        operator: rule.operator,
        operand1: rule.operand1,
        operand2: rule.operand2,
        listValues: rule.listValues,
        listSource: rule.listSource,
        inputMessage: rule.inputMessage,
        errorMessage: rule.errorMessage,
        response: rule.response,
      });
      if (rule.preserved) OdsPreserved.setValidation(validation, { element: rule.preserved });
      OdsPreserved.setValidationSource(validation, rule.source);
      sheet.addValidation(validation);
    }
  }
}
