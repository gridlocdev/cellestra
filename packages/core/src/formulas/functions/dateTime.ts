import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import { DateConversion, type CivilDateTime } from '../evaluation/dateConversion.ts';
import { defineFunction, type FormulaFunction } from './formulaFunction.ts';
import { Arg, createMetadata } from './functionMetadata.ts';
import { isFailure, readInt, readNumber } from './helpers.ts';

const outOfRange = (): CellValue =>
  CellValue.error(CellError.value('The date is out of the supported range.'));

function component(
  name: string,
  summary: string,
  example: string,
  select: (d: CivilDateTime) => number,
): FormulaFunction {
  return defineFunction(
    createMetadata(name, 'dateTime', summary, 'number', example, Arg.number('date')),
    (args, ev) => {
      const serial = readNumber(ev, args[0]!);
      if (isFailure(serial)) return serial.failed;
      const date = DateConversion.tryToDateTime(serial, true);
      return date ? CellValue.number(select(date)) : outOfRange();
    },
  );
}

/** The date & time functions. TODAY/NOW are volatile (recomputed each pass). */
export function dateTimeFunctions(): FormulaFunction[] {
  return [
    defineFunction(
      createMetadata('TODAY', 'dateTime', "Returns today's date.", 'number', 'TODAY()'),
      (_, ev) => {
        const now = ev.context.now();
        return CellValue.number(Math.floor(DateConversion.dateToSerial(now.year, now.month, now.day)));
      },
    ),
    defineFunction(
      createMetadata('NOW', 'dateTime', 'Returns the current date and time.', 'number', 'NOW()'),
      (_, ev) => CellValue.number(DateConversion.toSerial(ev.context.now())),
    ),
    defineFunction(
      createMetadata(
        'DATE',
        'dateTime',
        'Builds a date from year, month, and day.',
        'number',
        'DATE(2026, 6, 7)',
        Arg.number('year'),
        Arg.number('month'),
        Arg.number('day'),
      ),
      (args, ev) => {
        const year = readInt(ev, args[0]!);
        if (isFailure(year)) return year.failed;
        const month = readInt(ev, args[1]!);
        if (isFailure(month)) return month.failed;
        const day = readInt(ev, args[2]!);
        if (isFailure(day)) return day.failed;
        try {
          return CellValue.number(DateConversion.dateToSerial(year, month, day));
        } catch (e) {
          if (e instanceof RangeError) {
            return CellValue.error(CellError.value('DATE is out of the supported range.'));
          }
          throw e;
        }
      },
    ),
    component('YEAR', 'Returns the year of a date.', 'YEAR(A1)', (d) => d.year),
    component('MONTH', 'Returns the month (1–12) of a date.', 'MONTH(A1)', (d) => d.month),
    component('DAY', 'Returns the day of the month (1–31) of a date.', 'DAY(A1)', (d) => d.day),
    defineFunction(
      createMetadata(
        'WEEKDAY',
        'dateTime',
        'Returns the day of the week of a date as a number.',
        'number',
        'WEEKDAY(A1, 2)',
        Arg.number('date'),
        Arg.number('type', { optional: true }),
      ),
      (args, ev) => {
        const serial = readNumber(ev, args[0]!);
        if (isFailure(serial)) return serial.failed;
        let type = 1;
        if (args.length > 1) {
          const t = readInt(ev, args[1]!);
          if (isFailure(t)) return t.failed;
          type = t;
        }
        if (!DateConversion.tryToDateTime(serial, true)) return outOfRange();
        const dow = DateConversion.dayOfWeek(serial); // Sunday = 0 … Saturday = 6
        switch (type) {
          case 1:
            return CellValue.number(dow + 1); // Sunday = 1 … Saturday = 7
          case 2:
            return CellValue.number(((dow + 6) % 7) + 1); // Monday = 1 … Sunday = 7
          case 3:
            return CellValue.number((dow + 6) % 7); // Monday = 0 … Sunday = 6
          default:
            return CellValue.error(CellError.value('WEEKDAY type must be 1, 2, or 3.'));
        }
      },
    ),
    defineFunction(
      createMetadata(
        'NETWORKDAYS',
        'dateTime',
        'Counts the working days (Mon–Fri) between two dates.',
        'number',
        'NETWORKDAYS(A1, A2, Holidays)',
        Arg.number('start'),
        Arg.number('end'),
        Arg.range('holidays', { optional: true }),
      ),
      (args, ev) => {
        const startSerial = readNumber(ev, args[0]!);
        if (isFailure(startSerial)) return startSerial.failed;
        const endSerial = readNumber(ev, args[1]!);
        if (isFailure(endSerial)) return endSerial.failed;
        const holidays = new Set<number>();
        if (args.length > 2) {
          for (const value of ev.enumerateValues(args[2]!)) {
            if (value.kind === 'error') return value;
            if (value.kind === 'number') holidays.add(Math.floor(value.value));
          }
        }
        let start = Math.floor(startSerial);
        let end = Math.floor(endSerial);
        if (!DateConversion.tryToDateTime(start, true) || !DateConversion.tryToDateTime(end, true))
          return outOfRange();
        let sign = 1;
        if (start > end) {
          [start, end] = [end, start];
          sign = -1;
        }
        let count = 0;
        for (let day = start; day <= end; day++) {
          const dow = DateConversion.dayOfWeek(day);
          if (dow === 0 || dow === 6 || holidays.has(day)) continue;
          count++;
        }
        return CellValue.number(sign * count);
      },
    ),
  ];
}
