import { NumberFormatter } from '../../styling/numberFormatter.ts';
import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import { Coercion } from '../evaluation/coercion.ts';
import { safeLocale } from '../evaluation/localeInfo.ts';
import { tryParseNumberOrDate } from '../evaluation/valueParsing.ts';
import { defineFunction, type FormulaFunction } from './formulaFunction.ts';
import { Arg, createMetadata } from './functionMetadata.ts';
import { isFailure, readInt, readText } from './helpers.ts';

function edge(name: string, summary: string, example: string, fromLeft: boolean): FormulaFunction {
  return defineFunction(
    createMetadata(
      name,
      'text',
      summary,
      'text',
      example,
      Arg.text('text'),
      Arg.number('count', { optional: true }),
    ),
    (args, ev) => {
      const text = readText(ev, args[0]!);
      if (isFailure(text)) return text.failed;
      let count = 1;
      if (args.length > 1) {
        const c = readInt(ev, args[1]!);
        if (isFailure(c)) return c.failed;
        count = c;
      }
      if (count < 0) return CellValue.error(CellError.value(`${name} count must be ≥ 0.`));
      count = Math.min(count, text.length);
      return CellValue.text(fromLeft ? text.slice(0, count) : text.slice(text.length - count));
    },
  );
}

function transform(
  name: string,
  summary: string,
  example: string,
  op: (text: string, locale: string) => string,
): FormulaFunction {
  return defineFunction(
    createMetadata(name, 'text', summary, 'text', example, Arg.text('text')),
    (args, ev) => {
      const text = readText(ev, args[0]!);
      return isFailure(text) ? text.failed : CellValue.text(op(text, ev.context.locale));
    },
  );
}

function collapseSpaces(text: string): string {
  let out = '';
  let previousSpace = false;
  for (const c of text.trim()) {
    if (c === ' ') {
      if (!previousSpace) out += ' ';
      previousSpace = true;
    } else {
      out += c;
      previousSpace = false;
    }
  }
  return out;
}

function substitute(text: string, oldText: string, newText: string, instance: number | null): string {
  if (instance === null) return text.replaceAll(oldText, newText);
  let occurrence = 0;
  let index = 0;
  while ((index = text.indexOf(oldText, index)) >= 0) {
    occurrence++;
    if (occurrence === instance) return text.slice(0, index) + newText + text.slice(index + oldText.length);
    index += oldText.length;
  }
  return text;
}

/** The text functions. CONCAT also backs the `&` operator's library entry. */
export function textFunctions(): FormulaFunction[] {
  return [
    edge('LEFT', 'Returns the first characters of a string.', 'LEFT("Cellestra", 4)', true),
    edge('RIGHT', 'Returns the last characters of a string.', 'RIGHT("Cellestra", 3)', false),

    defineFunction(
      createMetadata(
        'MID',
        'text',
        'Returns characters from the middle of a string.',
        'text',
        'MID("Cellestra", 2, 4)',
        Arg.text('text'),
        Arg.number('start'),
        Arg.number('length'),
      ),
      (args, ev) => {
        const text = readText(ev, args[0]!);
        if (isFailure(text)) return text.failed;
        const start = readInt(ev, args[1]!);
        if (isFailure(start)) return start.failed;
        const length = readInt(ev, args[2]!);
        if (isFailure(length)) return length.failed;
        if (start < 1 || length < 0) {
          return CellValue.error(CellError.value('MID start must be ≥ 1 and length ≥ 0.'));
        }
        if (start > text.length) return CellValue.text('');
        const available = Math.min(length, text.length - (start - 1));
        return CellValue.text(text.substring(start - 1, start - 1 + available));
      },
    ),

    defineFunction(
      createMetadata(
        'LEN',
        'text',
        'Returns the number of characters in a string.',
        'number',
        'LEN("Cellestra")',
        Arg.text('text'),
      ),
      (args, ev) => {
        const text = readText(ev, args[0]!);
        return isFailure(text) ? text.failed : CellValue.number(text.length);
      },
    ),

    transform('LOWER', 'Converts text to lowercase.', 'LOWER("ABC")', (t, l) =>
      t.toLocaleLowerCase(safeLocale(l)),
    ),
    transform('UPPER', 'Converts text to uppercase.', 'UPPER("abc")', (t, l) =>
      t.toLocaleUpperCase(safeLocale(l)),
    ),
    transform('TRIM', 'Removes extra spaces, leaving single spaces between words.', 'TRIM("  a  b ")', (t) =>
      collapseSpaces(t),
    ),

    defineFunction(
      createMetadata(
        'SUBSTITUTE',
        'text',
        'Replaces occurrences of old text with new text.',
        'text',
        'SUBSTITUTE("a-b-c", "-", "+")',
        Arg.text('text'),
        Arg.text('old_text'),
        Arg.text('new_text'),
        Arg.number('instance', { optional: true }),
      ),
      (args, ev) => {
        const text = readText(ev, args[0]!);
        if (isFailure(text)) return text.failed;
        const oldText = readText(ev, args[1]!);
        if (isFailure(oldText)) return oldText.failed;
        const newText = readText(ev, args[2]!);
        if (isFailure(newText)) return newText.failed;
        if (oldText.length === 0) return CellValue.text(text);
        let instance: number | null = null;
        if (args.length > 3) {
          const parsed = readInt(ev, args[3]!);
          if (isFailure(parsed)) return parsed.failed;
          if (parsed < 1) return CellValue.error(CellError.value('SUBSTITUTE instance must be ≥ 1.'));
          instance = parsed;
        }
        return CellValue.text(substitute(text, oldText, newText, instance));
      },
    ),

    defineFunction(
      createMetadata(
        'VALUE',
        'text',
        'Converts text that represents a number into a number.',
        'number',
        'VALUE("123")',
        Arg.text('text'),
      ),
      (args, ev) => {
        const text = readText(ev, args[0]!);
        if (isFailure(text)) return text.failed;
        const n = tryParseNumberOrDate(text, ev.context.locale);
        return n === null
          ? CellValue.error(CellError.value(`'${text}' is not a number.`))
          : CellValue.number(n);
      },
    ),

    defineFunction(
      createMetadata(
        'TEXT',
        'text',
        'Formats a number or date as text using a format code.',
        'text',
        'TEXT(0.25, "0.0%")',
        Arg.any('value'),
        Arg.text('format'),
      ),
      (args, ev) => {
        const value = ev.evaluate(args[0]!);
        if (value.kind === 'error') return value;
        const format = readText(ev, args[1]!);
        if (isFailure(format)) return format.failed;
        // One formatting path shared with the grid and status bar.
        return CellValue.text(NumberFormatter.format(value, format, ev.context.locale));
      },
    ),

    defineFunction(
      createMetadata(
        'CONCAT',
        'text',
        'Joins several text values into one.',
        'text',
        'CONCAT(A1, " ", B1)',
        Arg.any('text', { repeatable: true }),
      ),
      (args, ev) => {
        let out = '';
        for (const arg of args) {
          for (const value of ev.enumerateValues(arg)) {
            if (value.kind === 'error') return value;
            const t = Coercion.toText(value);
            if (typeof t === 'string') out += t;
          }
        }
        return CellValue.text(out);
      },
    ),
  ];
}
