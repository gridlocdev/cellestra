import {
  chmodSync,
  linkSync,
  lstatSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  CellValue,
  OdsPackage,
  readOds,
  readOdsWithPackage,
  Workbook,
  writeOds,
  type Worksheet,
} from '@cellestra/core';
import { runCli } from '../src/cli.ts';
import { buildRoot } from '../src/index.ts';
import { isExpectedError } from '../src/infra/errors.ts';
import type { CliIo } from '../src/infra/io.ts';

/**
 * Drives the CLI in-process: every case gets a fresh temp folder, the
 * streams are captured, and the app launcher is a stub that never opens a window.
 */

interface Result {
  readonly exit: number;
  readonly out: string;
  readonly err: string;
}

let dir = '';
const launched: string[] = [];

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'cellestra-cli-'));
  launched.length = 0;
});

afterEach(() => {
  rmSync(dir, { recursive: true, force: true });
});

const path2 = (name: string): string => join(dir, name);

function run(...args: string[]): Result {
  let out = '';
  let err = '';
  const io: CliIo = {
    stdout: (text) => {
      out += text;
    },
    stderr: (text) => {
      err += text;
    },
    stdoutBytes: (bytes) => {
      out += new TextDecoder().decode(bytes);
    },
    readStdin: () => new Uint8Array(),
    homeDir: () => dir,
    launch: (path) => {
      launched.push(path);
      return 0;
    },
  };
  const exit = runCli(args, io);
  return { exit, out, err };
}

/** A parsed report, loosely typed so a test can read nested properties directly. */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Json = any;
const json = (text: string): Json => JSON.parse(text);
const read = (path: string): Workbook => readOds(readFileSync(path));
const sheet1 = (path: string, name = 'Sheet1'): Worksheet => read(path).findSheet(name)!;

describe('basics', () => {
  it('eval computes a formula', () => {
    const r = run('eval', '=ROUND(355/113, 4)', '--locale', 'en-US');
    expect(r.exit).toBe(0);
    expect(r.out.trim()).toBe('3.1416');
  });

  it('convert csv to ods round-trips a value', () => {
    const csv = path2('in.csv');
    const ods = path2('out.ods');
    writeFileSync(csv, 'Name,Score\nAda,90\nBev,75\n');
    expect(run('convert', csv, ods, '--sheet', 'Imported').exit).toBe(0);
    const r = run('get', ods, 'Imported.B2');
    expect(r.exit).toBe(0);
    expect(r.out.trim()).toBe('90');
  });

  it('set writes a formula and get reads back value and source', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run('set', ods, 'Sheet1.A1', '10', '--in-place');
    run('set', ods, 'Sheet1.A2', '5', '--in-place');
    expect(run('set', ods, 'Sheet1.A3', '=A1+A2', '--in-place').exit).toBe(0);
    expect(run('get', ods, 'Sheet1.A3').out.trim()).toBe('15');
    expect(run('get', ods, 'Sheet1.A3', '--formula').out.trim()).toBe('=A1+A2');
  });

  it('set refuses to write without --in-place or --output', () => {
    const ods = path2('book.ods');
    run('new', ods);
    const r = run('set', ods, 'Sheet1.A1', '1');
    expect(r.exit).toBe(1);
    expect(r.err).toContain('--in-place');
  });

  it('validate passes clean and fails on errors', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run('set', ods, 'Sheet1.A1', '1', '--in-place');
    expect(run('validate', ods).exit).toBe(0);
    run('set', ods, 'Sheet1.A2', '=1/0', '--in-place');
    const r = run('validate', ods);
    expect(r.exit).toBe(1);
    expect(r.err).toContain('#DIV/0!');
  });

  it('functions lists and describes', () => {
    expect(run('functions').out).toContain('SUM');
    const r = run('functions', 'VLOOKUP');
    expect(r.exit).toBe(0);
    expect(r.out).toContain('VLOOKUP(');
    expect(r.out).toContain('Example:');
  });

  it('diff exits nonzero when workbooks differ', () => {
    const a = path2('a.ods');
    const b = path2('b.ods');
    run('new', a);
    run('set', a, 'Sheet1.A1', '1', '--in-place');
    run('convert', a, b);
    expect(run('diff', a, b).exit).toBe(0);
    run('set', b, 'Sheet1.A1', '2', '--in-place');
    const r = run('diff', a, b);
    expect(r.exit).toBe(1);
    expect(r.out).toContain('Sheet1.A1');
  });
});

describe('--json reports', () => {
  it('get --json reports a typed value and the formula', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run('set', ods, 'Sheet1.A1', '10', '--in-place');
    run('set', ods, 'Sheet1.A2', '=A1*2', '--in-place');
    const r = run('get', ods, 'Sheet1.A2', '--json');
    expect(r.exit).toBe(0);
    const doc = json(r.out);
    expect(doc.cell).toBe('Sheet1.A2');
    expect(doc.value).toBe(20);
    expect(doc.display).toBe('20');
    expect(doc.formula).toBe('=A1*2');
    expect(doc).not.toHaveProperty('error');
  });

  it('styling and layout read back through json', () => {
    const ods = path2('book.ods');
    const ops = path2('ops.json');
    run('new', ods);
    writeFileSync(
      ops,
      `[{"op":"set","cell":"A1","value":"Head"},
        {"op":"style","range":"A1","bold":true,"fill":"#EEEEEE","align":"center","border":"thin","format":"0.00"},
        {"op":"width","columns":"A","width":120},{"op":"height","rows":"1","height":30},
        {"op":"hide","columns":"C"},{"op":"merge","range":"B2:C3"},{"op":"freeze","rows":1}]`,
    );
    expect(run('apply', ods, ops, '--in-place').exit).toBe(0);

    const style = json(run('get', ods, 'A1', '--json').out).style;
    expect(style.bold).toBe(true);
    expect(style.fill).toBe('#EEEEEE');
    expect(style.align).toBe('center');
    expect(style.border).toBe('thin');
    expect(style).not.toHaveProperty('italic');

    expect(json(run('get', ods, 'A2', '--json').out)).not.toHaveProperty('style');

    const sheet = json(run('sheets', ods, '--json').out)[0];
    expect(sheet.widths.A).toBe(120);
    expect(sheet.heights['1']).toBe(30);
    expect(sheet.hiddenColumns[0]).toBe('C');
    expect(sheet.merges[0]).toBe('B2:C3');
    expect(sheet.frozenRows).toBe(1);
  });

  it('names and tables are addresses, and get and set take several', () => {
    const ods = path2('book.ods');
    run('new', ods);
    expect(
      run('set', ods, 'A1=Item', 'B1=Qty', 'A2=Pen', 'B2=3', 'A3=Ink', 'B3=4', 'B4==SUM(B2:B3)', '--in-place')
        .exit,
    ).toBe(0);
    run('name', 'add', ods, 'Total', 'Sheet1.B4', '--in-place');
    run('table', ods, 'A1:B3', '--name', 'Orders', '--in-place');

    expect(run('get', ods, 'Total').out.trim()).toBe('7');
    expect(json(run('cat', ods, '--range', 'Orders', '--json').out).range).toBe('A1:B3');

    const r = run('get', ods, 'A2', 'B2:B3');
    expect(r.exit).toBe(0);
    expect(r.out.trim().split('\n')).toEqual(['Sheet1.A2\tPen', 'Sheet1.B2\t3', 'Sheet1.B3\t4']);

    expect(json(run('get', ods, 'A1', 'Total', '--json').out)).toHaveLength(2);
    expect(run('get', ods, 'Orders').out.trim().split('\n')).toHaveLength(6);
  });

  it('trace follows precedents and dependents', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run('set', ods, 'A1=1', 'A2=2', 'A3=3', 'B1==SUM(A1:A3)', 'C1==B1*2', '--in-place');

    const r = run('trace', ods, 'C1', '--depth', '0');
    expect(r.exit).toBe(0);
    const lines = r.out.trim().split('\n');
    expect(lines[0]!.startsWith('Sheet1.C1  =B1*2  → 12')).toBe(true);
    expect(lines[1]).toContain('  Sheet1.B1  =SUM(A1:A3)  → 6');
    expect(lines[2]).toContain('Sheet1.A1:A3 (3 cells)');

    const tree = json(run('trace', ods, 'A2', '--dependents', '--depth', '0', '--json').out).tree;
    expect(tree.dependents[0].address).toBe('Sheet1.B1');
    expect(tree.dependents[0].dependents[0].address).toBe('Sheet1.C1');
  });

  it('lint finds mistakes that are not errors', () => {
    const ods = path2('book.ods');
    const ops = path2('ops.json');
    run('new', ods);
    writeFileSync(
      ops,
      `[{"op":"set","at":"A1","values":[["Item","Qty"],["a",1],["b",2],["c",3],["d","4"],["Sum","=SUM(B2:B4)"]]},
        {"op":"set","at":"D1","values":[["=A2"],["=A3"],["=A9"],["=A5"]]},
        {"op":"table","range":"A1:B5","name":"T"},
        {"op":"name","name":"Unused","range":"B2"},
        {"op":"validation","range":"B2:B5","list":["1","2"]}]`,
    );
    expect(run('apply', ods, ops, '--in-place').exit).toBe(0);

    // apply parses "4" like typing (a number), so plant a genuine text "4" the way a CSV import can.
    const { workbook, package: pkg } = readOdsWithPackage(readFileSync(ods));
    workbook.activeSheet.setValue(1, 2, CellValue.text('4')); // C2
    writeFileSync(ods, writeOds(workbook, pkg));

    const r = run('lint', ods, '--json', '--locale', 'en-US');
    expect(r.exit).toBe(0);
    const codes = (json(r.out).findings as { code: string; cell: string }[]).map(
      (f) => `${f.code}@${f.cell}`,
    );
    expect(codes).toContain('number-as-text@Sheet1.C2');
    expect(codes).toContain('short-range@Sheet1.B6');
    expect(codes).toContain('inconsistent-formula@Sheet1.D3');
    expect(codes).toContain('validation@Sheet1.B4');
    expect(codes).toContain('unused-name@Unused');
    expect(codes.some((c) => c.startsWith('error'))).toBe(false);

    expect(run('lint', ods, '--strict').exit).toBe(1);
  });

  it('put drops a csv grid at an anchor, parsed like typing', () => {
    const ods = path2('book.ods');
    const csv = path2('grid.csv');
    run('new', ods);
    writeFileSync(csv, 'Item,Qty\nPen,3\nInk,"=C3*2"\n');
    const r = run('put', ods, 'B2', '--from', csv, '--in-place', '--json');
    expect(r.exit).toBe(0);
    expect(json(r.out).range).toBe('B2:C4');
    expect(run('get', ods, 'C3').out.trim()).toBe('3');
    expect(run('get', ods, 'C4').out.trim()).toBe('6');
    expect(run('get', ods, 'C4', '--formula').out.trim()).toBe('=C3*2');
  });

  it('style --checkbox makes TRUE/FALSE tick boxes', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run('set', ods, 'A1=true', '--in-place');
    expect(run('style', ods, 'A1:A2', '--checkbox', '--in-place').exit).toBe(0);
    const a2 = json(run('get', ods, 'A2', '--json').out);
    expect(a2.value).toBe(false);
    expect(a2.style.checkbox).toBe(true);
    expect(json(run('get', ods, 'A1', '--json').out).value).toBe(true);
    expect(run('style', ods, 'A1:A2', '--checkbox', 'false', '--in-place').exit).toBe(0);
    expect(json(run('get', ods, 'A2', '--json').out)).not.toHaveProperty('style');
  });

  it('layout verbs front the apply ops', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run('set', ods, 'A1=Head', 'A2=1', '--in-place');
    expect(
      run('style', ods, 'A1', '--bold', '--fill', '#EEEEEE', '--align', 'center', '--in-place').exit,
    ).toBe(0);
    expect(run('style', ods, 'A1', '--bold', 'false', '--in-place').exit).toBe(0);
    expect(run('width', ods, 'A:B', '100', '--in-place').exit).toBe(0);
    expect(run('height', ods, '1', '30', '--in-place').exit).toBe(0);
    expect(run('freeze', ods, '--rows', '1', '--in-place').exit).toBe(0);
    expect(run('settings', ods, '--locale', 'de-DE', '--calc', 'manual', '--in-place').exit).toBe(0);
    expect(run('add-sheet', ods, 'Data', '--at', '1', '--in-place').exit).toBe(0);
    expect(run('add-sheet', ods, 'Data', '--if-missing', '--in-place').exit).toBe(0);
    expect(run('duplicate-sheet', ods, 'Sheet1', '--as', 'Copy', '--in-place').exit).toBe(0);
    expect(run('move-sheet', ods, 'Copy', '1', '--in-place').exit).toBe(0);
    expect(run('clear', ods, 'Sheet1.A2', '--in-place').exit).toBe(0);

    const style = json(run('get', ods, 'Sheet1.A1', '--json').out).style;
    expect(style).not.toHaveProperty('bold');
    expect(style.fill).toBe('#EEEEEE');
    const info = json(run('info', ods, '--json').out);
    expect(info.locale).toBe('de-DE');
    expect(info.calculation).toBe('Manual');
    expect((info.sheets as { name: string }[]).map((s) => s.name)).toEqual(['Copy', 'Data', 'Sheet1']);
    expect(info.sheets[2].widths.B).toBe(100);
    expect(run('get', ods, 'Sheet1.A2').out.trim()).toBe('');
  });

  it('copy, move, fill, and paint behave like the app', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run('set', ods, 'A1=1', 'A2=2', 'A3=3', 'B1==A1*10', '--in-place');
    run('style', ods, 'B1', '--bold', '--in-place');

    expect(run('fill', ods, 'B1:B3', '--in-place').exit).toBe(0);
    expect(run('get', ods, 'B2', '--formula').out.trim()).toBe('=A2*10');
    expect(run('get', ods, 'B3').out.trim()).toBe('30');

    expect(run('copy', ods, 'A1:B3', 'D1', '--in-place').exit).toBe(0);
    expect(run('get', ods, 'E2', '--formula').out.trim()).toBe('=D2*10');
    expect(run('copy', ods, 'A1:B3', 'G1', '--mode', 'values', '--in-place').exit).toBe(0);
    expect(run('get', ods, 'H1').out.trim()).toBe('10');
    expect(run('get', ods, 'H1', '--formula').out.trim()).toBe('10');

    expect(run('move', ods, 'D1:E3', 'Sheet1.D10', '--in-place').exit).toBe(0);
    expect(run('get', ods, 'D1').out.trim()).toBe('');
    expect(run('get', ods, 'E11', '--formula').out.trim()).toBe('=D11*10');

    expect(run('paint', ods, 'B1', 'A1:A3', '--in-place').exit).toBe(0);
    const doc = json(run('get', ods, 'A3', '--json').out);
    expect(doc.style.bold).toBe(true);
    expect(doc.display).toBe('3');
  });

  it('filter hides rows and --clear shows them again', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run('set', ods, 'A1=Name', 'B1=Score', 'A2=Ann', 'B2=5', 'A3=Bob', 'B3=9', 'A4=Cy', 'B4=7', '--in-place');

    const r = run('filter', ods, 'A1:B4', '--where', 'B >6', '--values', 'A Bob,Cy', '--in-place', '--json');
    expect(r.exit).toBe(0);
    expect(json(r.out).hiddenRows).toBe(1);
    expect(json(run('sheets', ods, '--json').out)[0].hiddenRows[0]).toBe('2');

    expect(run('filter', ods, 'A1:B4', '--clear', '--in-place').exit).toBe(0);
    expect(json(run('sheets', ods, '--json').out)[0]).not.toHaveProperty('hiddenRows');
  });

  it('table options and removal, and validation kinds', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run('set', ods, 'A1=Name', 'B1=Age', 'A2=Ann', 'B2=30', 'A3=Bob', 'B3=41', '--in-place');
    expect(run('table', ods, 'A1:B3', '--name', 'People', '--in-place').exit).toBe(0);
    expect(
      run(
        'table',
        ods,
        'People',
        '--banded',
        'true',
        '--header-column',
        'true',
        '--resize',
        'A1:B4',
        '--in-place',
      ).exit,
    ).toBe(0);
    const table = json(run('sheets', ods, '--json').out)[0].tables[0];
    expect(table.range).toBe('A1:B4');
    expect(table.bandedRows).toBe(true);
    expect(table.headerColumn).toBe(true);

    expect(
      run(
        'data-validation',
        ods,
        'B2:B4',
        '--kind',
        'whole',
        '--op',
        'between',
        '--value',
        '18',
        '--max',
        '65',
        '--error',
        'Adults only',
        '--in-place',
      ).exit,
    ).toBe(0);
    expect(
      run(
        'data-validation',
        ods,
        'A2:A4',
        '--kind',
        'text-length',
        '--op',
        'less',
        '--value',
        '10',
        '--response',
        'warning',
        '--in-place',
      ).exit,
    ).toBe(0);
    const rules = Object.fromEntries(
      (json(run('sheets', ods, '--json').out)[0].validations as { range: string }[]).map((r) => [r.range, r]),
    ) as Record<string, Json>;
    expect(Object.keys(rules)).toHaveLength(2);
    expect(rules['B2:B4'].kind).toBe('wholenumber');
    expect(rules['B2:B4'].value2).toBe(65);
    expect(rules['A2:A4'].response).toBe('warning');

    // A Stop rule refuses a fill that would break it, as in the app; lint reports an existing violation.
    run('set', ods, 'B2=99', '--in-place');
    const findings = json(run('lint', ods, '--json').out).findings as { code: string; cell: string }[];
    expect(findings.some((f) => f.code === 'validation' && f.cell === 'Sheet1.B2')).toBe(true);
    expect(run('fill', ods, 'B2:B4', '--in-place').exit).toBe(1);

    expect(run('data-validation', ods, 'B2:B4', '--remove', '--in-place').exit).toBe(0);
    expect(run('table', ods, 'People', '--remove', '--in-place').exit).toBe(0);
    const after = json(run('sheets', ods, '--json').out)[0];
    expect(after).not.toHaveProperty('tables');
    expect(after.validations).toHaveLength(1);
  });

  it('find --replace, stats, cat --both, and a structural diff', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run(
      'set',
      ods,
      'A1=apple pie',
      'A2=Apple tart',
      'A3=pear',
      'B1=1',
      'B2=2',
      'B3=4',
      'B4==SUM(B1:B3)',
      '--in-place',
    );

    let r = run('find', ods, 'apple', '--replace', 'plum', '--in-place', '--json');
    expect(r.exit).toBe(0);
    expect(json(r.out).replaced).toBe(2);
    expect(run('get', ods, 'A2').out.trim()).toBe('plum tart');
    expect(run('find', ods, 'B1:B3', '--in', 'formulas', '--replace', 'B1:B2', '--in-place').exit).toBe(0);
    expect(run('get', ods, 'B4', '--formula').out.trim()).toBe('=SUM(B1:B2)');

    const stats = json(run('stats', ods, 'A1:B4', '--json').out);
    expect(stats.cells).toBe(8);
    expect(stats.blanks).toBe(1);
    expect(stats.types.text).toBe(3);
    expect(stats.median).toBe(2.5);
    expect(stats.distinct).toBe(7);

    expect(run('cat', ods, '--both').out).toContain('=SUM(B1:B2) → 3');

    const other = path2('other.ods');
    run('convert', ods, other);
    run('freeze', other, '--rows', '1', '--in-place');
    run('style', other, 'A1', '--bold', '--in-place');
    run('name', 'add', other, 'Total', 'Sheet1.B4', '--in-place');
    expect(run('diff', ods, ods).exit).toBe(0);
    r = run('diff', ods, other, '--styles', '--json');
    expect(r.exit).toBe(1);
    const doc = json(r.out);
    const structure = (doc.structure as { property: string }[]).map((d) => d.property);
    expect(structure).toContain('frozenRows');
    expect(structure).toContain('namedRanges');
    expect(
      (doc.cells as { styleAfter?: string }[]).some(
        (c) => typeof c.styleAfter === 'string' && c.styleAfter.includes('bold'),
      ),
    ).toBe(true);

    expect(run('diff', ods, other).exit).toBe(1);

    expect(run('find', ods, '=SUM(', '--in', 'formulas', '--replace', '=AVERAGE(', '--in-place').exit).toBe(
      0,
    );
    expect(run('get', ods, 'B4', '--formula').out.trim()).toBe('=AVERAGE(B1:B2)');
  });

  it('completions cover subcommands and options', () => {
    const zsh = run('completions', 'zsh');
    expect(zsh.exit).toBe(0);
    expect(zsh.out).toContain("'set') opts=(");
    expect(zsh.out).toContain("'--in-place[Edit the file in place (atomic write)]'");
    expect(zsh.out).toContain("'name add') opts=(");

    const bash = run('completions', 'bash');
    expect(bash.exit).toBe(0);
    expect(bash.out).toContain('apply) opts="');
    expect(bash.out).toContain('--dry-run');
    expect(bash.out).toContain('charts)');
  });

  it('set --json reports the changed cell and any error it introduced', () => {
    const ods = path2('book.ods');
    run('new', ods);
    const r = run('set', ods, 'Sheet1.B1', '=SUMM(A1)', '--in-place', '--json');
    expect(r.exit).toBe(0);
    expect(r.err).toBe('');
    const doc = json(r.out);
    expect(doc.ok).toBe(true);
    expect(doc.changed).toHaveLength(1);
    expect(doc.changed[0].cell).toBe('Sheet1.B1');
    expect(doc.changed[0].error).toBe('#NAME?');
    expect(doc.errors[0].cell).toBe('Sheet1.B1');
  });

  it('validate --json lists errors and keeps the exit code', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run('set', ods, 'Sheet1.A1', '=1/0', '--in-place');
    const r = run('validate', ods, '--json');
    expect(r.exit).toBe(1);
    const doc = json(r.out);
    expect(doc.ok).toBe(false);
    expect(doc.errors[0].error).toBe('#DIV/0!');
  });

  it('cat --json carries the grid and every populated cell', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run('set', ods, 'Sheet1.A1', 'Name', '--in-place');
    run('set', ods, 'Sheet1.B2', '3', '--in-place');
    const r = run('cat', ods, '--json');
    expect(r.exit).toBe(0);
    const doc = json(r.out);
    expect(doc.range).toBe('A1:B2');
    expect(doc.rows).toHaveLength(2);
    expect(doc.cells).toHaveLength(2);
    expect(doc.cells[0].value).toBe('Name');
    expect(doc.cells[1].value).toBe(3);
  });

  it('diff --json is structured and exits nonzero on a difference', () => {
    const a = path2('a.ods');
    const b = path2('b.ods');
    run('new', a);
    run('set', a, 'Sheet1.A1', '1', '--in-place');
    run('convert', a, b);
    run('set', b, 'Sheet1.A1', '2', '--in-place');
    const r = run('diff', a, b, '--json');
    expect(r.exit).toBe(1);
    const doc = json(r.out);
    expect(doc.same).toBe(false);
    expect(doc.cells[0].cell).toBe('Sheet1.A1');
    expect(doc.cells[0].before).toBe('1');
    expect(doc.cells[0].after).toBe('2');
  });

  it('eval and functions speak json', () => {
    let r = run('eval', '=1+1', '--json');
    expect(r.exit).toBe(0);
    expect(json(r.out).value).toBe(2);
    r = run('functions', 'ABS', '--json');
    expect(r.exit).toBe(0);
    const doc = json(r.out);
    expect(doc.name).toBe('ABS');
    expect(doc.arguments[0].name).toBe('number');
  });

  it('a user error under --json is a json document too', () => {
    const ods = path2('book.ods');
    run('new', ods);
    const r = run('set', ods, 'Nope.A1', '1', '--in-place', '--json');
    expect(r.exit).toBe(1);
    expect(r.err.startsWith('cellestra: ')).toBe(true);
    const doc = json(r.out);
    expect(doc.ok).toBe(false);
    expect(doc.error).toContain('Nope');
  });

  it('a parse error under --json is a json document too', () => {
    const r = run('set', 'book.ods', '--json');
    expect(r.exit).toBe(1);
    const doc = json(r.out);
    expect(doc.ok).toBe(false);
    expect(doc.error.length).toBeGreaterThan(0);
  });
});

describe('apply', () => {
  it('runs a batch of edits', () => {
    const ods = path2('book.ods');
    const edits = path2('edits.json');
    run('new', ods);
    writeFileSync(edits, '[{"cell":"Sheet1.A1","value":"=2+3"},{"cell":"Sheet1.A2","value":"hi"}]');
    expect(run('apply', ods, edits, '--in-place').exit).toBe(0);
    expect(run('get', ods, 'Sheet1.A1').out.trim()).toBe('5');
    expect(run('get', ods, 'Sheet1.A2').out.trim()).toBe('hi');
  });

  it('builds a sheet from typed ops in one pass', () => {
    const ods = path2('book.ods');
    const ops = path2('ops.json');
    run('new', ods);
    writeFileSync(
      ops,
      `{"ops": [
        {"op":"set", "at":"A1", "values":[["Item","Qty","Price","Total"],["Pen",3,1.5,"=B2*C2"],["Ink",2,4,"=B3*C3"]]},
        {"op":"set", "cell":"D4", "value":"=SUM(D2:D3)"},
        {"op":"style", "range":"A1:D1", "bold":true, "fill":"#EEEEEE", "align":"center"},
        {"op":"style", "range":"C2:D4", "format":"0.00"},
        {"op":"width", "columns":"A", "width":120},
        {"op":"freeze", "rows":1},
        {"op":"add-sheet", "name":"Notes"},
        {"op":"name", "name":"GrandTotal", "range":"Sheet1.D4"},
        {"op":"table", "range":"A1:D3", "name":"Orders"}
      ]}`,
    );
    const r = run('apply', ods, ops, '--in-place', '--json');
    expect(r.exit).toBe(0);
    const doc = json(r.out);
    expect(doc.ok).toBe(true);
    expect(doc.applied).toBe(9);
    expect(doc.errors).toHaveLength(0);
    expect(
      (doc.changed as { cell: string; value: unknown }[]).some(
        (c) => c.cell === 'Sheet1.D4' && c.value === 12.5,
      ),
    ).toBe(true);

    const workbook = read(ods);
    const sheet = workbook.findSheet('Sheet1')!;
    expect(CellValue.asNumber(sheet.getValue(3, 3))).toBe(12.5);
    expect(workbook.styles.get(sheet.find(0, 0)!.styleId).font.bold).toBe(true);
    expect(workbook.styles.get(sheet.find(1, 2)!.styleId).numberFormat.code).toBe('0.00');
    expect(sheet.frozenRows).toBe(1);
    expect(workbook.findSheet('Notes')).toBeDefined();
    expect(workbook.findNamedRange('GrandTotal')).toBeDefined();
    expect(sheet.tables[0]!.name).toBe('Orders');
    expect(json(run('sheets', ods, '--json').out)[0].widths.A).toBe(120);
    expect(run('get', ods, 'Sheet1.D4').out.trim()).toBe('12.50');
  });

  it('accepts a bare array and the original cell/value form', () => {
    const ods = path2('book.ods');
    const ops = path2('ops.json');
    run('new', ods);
    writeFileSync(
      ops,
      '[{"cell":"A1","value":"7"},{"cell":"A2","formula":"A1*2"},{"op":"set","cell":"A3","value":true}]',
    );
    expect(run('apply', ods, ops, '--in-place').exit).toBe(0);
    expect(run('get', ods, 'A2').out.trim()).toBe('14');
    expect(run('get', ods, 'A2', '--formula').out.trim()).toBe('=A1*2');
    expect(run('get', ods, 'A3').out.trim()).toBe('TRUE');
  });

  it('--dry-run reports the outcome and writes nothing', () => {
    const ods = path2('book.ods');
    const ops = path2('ops.json');
    run('new', ods);
    writeFileSync(ops, '[{"op":"set","cell":"A1","value":"=NOPE(1)"}]');
    const before = readFileSync(ods);
    const r = run('apply', ods, ops, '--dry-run', '--json');
    expect(r.exit).toBe(0);
    const doc = json(r.out);
    expect(doc.dryRun).toBe(true);
    expect(doc.errors[0].error).toBe('#NAME?');
    expect(readFileSync(ods).equals(before)).toBe(true);
  });

  it('aborts the whole batch on a bad op and says which', () => {
    const ods = path2('book.ods');
    const ops = path2('ops.json');
    run('new', ods);
    writeFileSync(ops, '[{"op":"set","cell":"A1","value":1},{"op":"style","range":"A1","fill":"red"}]');
    const r = run('apply', ods, ops, '--in-place');
    expect(r.exit).toBe(1);
    expect(r.err).toContain('op 2 (style)');
    expect(run('get', ods, 'A1').out.trim()).toBe('');
  });

  it('names the op when core refuses', () => {
    const ods = path2('book.ods');
    const ops = path2('ops.json');
    run('new', ods);
    writeFileSync(ops, '[{"op":"add-sheet","name":"Sheet1"}]');
    const r = run('apply', ods, ops, '--in-place');
    expect(r.exit).toBe(1);
    expect(r.err).toContain('op 1 (add-sheet)');
  });

  it('apply and sheet verbs refuse to overwrite without being told to', () => {
    const ods = path2('book.ods');
    const ops = path2('ops.json');
    run('new', ods);
    writeFileSync(ops, '[]');
    expect(run('apply', ods, ops).err).toContain('--in-place');
    expect(run('add-sheet', ods, 'Two').err).toContain('--in-place');
  });
});

describe('reference', () => {
  it('man is one page generated from the live command tree', () => {
    let r = run('man');
    expect(r.exit).toBe(0);
    const root = buildRoot('0.0.0');
    for (const command of (root.subcommands ?? []).filter((c) => (c.subcommands ?? []).length === 0)) {
      expect(r.out).toContain(`### cellestra ${command.name}`);
    }
    expect(r.out).toContain('### cellestra name add');
    expect(r.out).toContain('--in-place');
    expect(r.out).toContain('[default: table]');
    expect(r.out).toContain('## Exit codes');
    expect(r.out).toContain('{"op":"freeze"');

    r = run('man', '--json');
    expect(r.exit).toBe(0);
    const doc = json(r.out);
    expect((doc.commands as { name: string }[]).some((c) => c.name === 'cellestra apply')).toBe(true);
  });

  it('skill prints the embedded skill and installs it into a skills directory', () => {
    const r = run('skill');
    expect(r.exit).toBe(0);
    expect(r.out.startsWith('---\nname: cellestra\n')).toBe(true);
    expect(r.out).toContain('cellestra apply');
    expect(r.out).not.toContain('AGENTS.md');

    const skills = path2('skills');
    const installed = run('skill', '--install', '--dir', skills);
    expect(installed.exit).toBe(0);
    const file = join(skills, 'cellestra', 'SKILL.md');
    expect(installed.err).toContain(file);
    expect(readFileSync(file, 'utf8')).toBe(r.out);
  });

  it('schema prints a contract and rejects an unknown one', () => {
    let r = run('schema', 'apply');
    expect(r.exit).toBe(0);
    expect(r.out).toContain('"op":"style"');
    r = run('schema', 'report');
    expect(r.exit).toBe(0);
    expect(r.out).toContain('"changed"');
    expect(run('schema', 'nope').exit).toBe(1);
  });
});

describe('robustness', () => {
  it('edits preserve unmodeled content', () => {
    const ods = path2('withchart.ods');
    const workbook = Workbook.createNew();
    workbook.activeSheet.setValue(0, 0, CellValue.number(1));
    const pkg = OdsPackage.read(writeOds(workbook));
    pkg.setBytes('Pictures/chart.bin', new Uint8Array([1, 2, 3, 4]));
    writeFileSync(ods, pkg.write());

    expect(run('set', ods, 'Sheet1.A1', '42', '--in-place').exit).toBe(0);

    const after = OdsPackage.read(readFileSync(ods));
    expect(after.contains('Pictures/chart.bin')).toBe(true);
    expect([...after.getBytes('Pictures/chart.bin')!]).toEqual([1, 2, 3, 4]);
  });

  it('stats aggregates a range', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run('set', ods, 'Sheet1.A1', '10', '--in-place');
    run('set', ods, 'Sheet1.A2', '20', '--in-place');
    run('set', ods, 'Sheet1.A3', '30', '--in-place');
    const r = run('stats', ods, 'Sheet1.A1:A3');
    expect(r.exit).toBe(0);
    expect(r.out).toContain('sum\t60');
    expect(r.out).toContain('average\t20');
    expect(r.out).toContain('max\t30');
  });

  it.each([
    [['get', '/no/such/file.ods', 'A1']],
    [['add-sheet', null, 'Sheet1', '--in-place']],
    [['rename-sheet', null, 'Sheet1', 'bad/name', '--in-place']],
  ] as (string | null)[][][])('user errors report cleanly and exit nonzero: %j', (template) => {
    const ods = path2('book.ods');
    run('new', ods);
    const r = run(...template.map((a) => a ?? ods));
    expect(r.exit).toBe(1);
    expect(r.err.trim().startsWith('cellestra:')).toBe(true);
    expect(r.err).not.toContain('Unhandled exception');
    expect(r.err).not.toContain('    at ');
  });

  it('an invalid regex reports cleanly', () => {
    const ods = path2('book.ods');
    run('new', ods);
    const r = run('find', ods, '[', '--regex');
    expect(r.exit).toBe(1);
    expect(r.err).toContain('Invalid regex');
    expect(r.err).not.toContain('    at ');
  });

  it('stats ignores empty cells in a large range', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run('set', ods, 'Sheet1.A1', '10', '--in-place');
    run('set', ods, 'Sheet1.A500', '40', '--in-place');
    const r = run('stats', ods, 'Sheet1.A1:A1000');
    expect(r.exit).toBe(0);
    expect(r.out).toContain('count\t2');
    expect(r.out).toContain('sum\t50');
    expect(r.out).toContain('min\t10');
    expect(r.out).toContain('max\t40');
  });

  it('cat --format csv is pipe friendly', () => {
    const ods = path2('book.ods');
    run('new', ods);
    run('set', ods, 'Sheet1.A1', 'x', '--in-place');
    run('set', ods, 'Sheet1.B1', 'y', '--in-place');
    const r = run('cat', ods, '--range', 'A1:B1', '--format', 'csv');
    expect(r.exit).toBe(0);
    expect(r.out.trim()).toBe('x,y');
  });
});

describe('data, structural, and layout edits', () => {
  function threeRowBook(): string {
    const ods = path2('data.ods');
    run('new', ods);
    run('set', ods, 'A1', 'Name', '--in-place');
    run('set', ods, 'B1', 'Score', '--in-place');
    run('set', ods, 'A2', 'Beta', '--in-place');
    run('set', ods, 'B2', '2', '--in-place');
    run('set', ods, 'A3', 'Alpha', '--in-place');
    run('set', ods, 'B3', '1', '--in-place');
    return ods;
  }

  function messyBook(): string {
    const csv = path2('messy.csv');
    const ods = path2('messy.ods');
    writeFileSync(csv, 'Region,Revenue\nWest,  1200.50 \nEast,900\n');
    run('import', csv, '--into', ods, '--sheet', 'Sheet1');
    return ods;
  }

  it('sort orders rows by a column, keeping the header', () => {
    const ods = threeRowBook();
    expect(run('sort', ods, 'A1:B3', '--by', 'A', '--in-place').exit).toBe(0);
    expect(run('get', ods, 'A1').out.trim()).toBe('Name');
    expect(run('get', ods, 'A2').out.trim()).toBe('Alpha');
    expect(run('get', ods, 'A3').out.trim()).toBe('Beta');
  });

  it('sort descending reverses a numeric column', () => {
    const ods = threeRowBook();
    run('sort', ods, 'A1:B3', '--by', 'B:desc', '--in-place');
    expect(run('get', ods, 'B2').out.trim()).toBe('2');
    expect(run('get', ods, 'B3').out.trim()).toBe('1');
  });

  it('remove-duplicates drops repeated rows and reports the count', () => {
    const ods = threeRowBook();
    run('set', ods, 'A4', 'Beta', '--in-place');
    run('set', ods, 'B4', '2', '--in-place');
    const r = run('remove-duplicates', ods, 'A1:B4', '--in-place');
    expect(r.exit).toBe(0);
    expect(r.err).toContain('Removed 1');
    expect(run('get', ods, 'A4').out.trim()).toBe('');
  });

  it('clean-up --dry-run reports without writing', () => {
    const ods = messyBook();
    const before = statSync(ods).mtimeMs;
    const r = run('clean-up', ods, '--dry-run', '--locale', 'en-US');
    expect(r.exit).toBe(0);
    expect(r.out).toContain('Sheet1:');
    expect(statSync(ods).mtimeMs).toBe(before);
  });

  it('clean-up refuses to overwrite without being told to', () => {
    const ods = messyBook();
    const r = run('clean-up', ods);
    expect(r.exit).not.toBe(0);
    expect(r.err).toContain('--in-place');
  });

  it('clean-up formats the workbook and settles after one pass', () => {
    const ods = messyBook();
    expect(run('clean-up', ods, '--in-place', '--locale', 'en-US').exit).toBe(0);
    const sheet = sheet1(ods);
    expect(CellValue.asNumber(sheet.getValue(1, 1))).toBe(1200.5);
    expect(sheet.frozenRows).toBe(1);
    expect(run('clean-up', ods, '--dry-run', '--locale', 'en-US').out).toContain('Nothing to clean up.');
  });

  it('clean-up can be narrowed to one sheet', () => {
    const ods = messyBook();
    run('add-sheet', ods, 'Other', '--in-place');
    run('set', ods, 'Other.A1', 'Label', '--in-place');
    run('set', ods, 'Other.A2', '7', '--in-place');
    const r = run('clean-up', ods, '--sheet', 'Sheet1', '--dry-run', '--locale', 'en-US');
    expect(r.exit).toBe(0);
    expect(r.out).not.toContain('Other:');
  });

  it('split separates a delimited column', () => {
    const ods = path2('split.ods');
    run('new', ods);
    run('set', ods, 'A1', 'a;b;c', '--in-place');
    expect(run('split', ods, 'A1:A1', '-d', ';', '--in-place').exit).toBe(0);
    expect(run('get', ods, 'A1').out.trim()).toBe('a');
    expect(run('get', ods, 'B1').out.trim()).toBe('b');
    expect(run('get', ods, 'C1').out.trim()).toBe('c');
  });

  it('table creation labels empty headers and round-trips', () => {
    const ods = path2('table.ods');
    run('new', ods);
    run('set', ods, 'A2', '1', '--in-place');
    expect(run('table', ods, 'A1:B2', '--in-place').exit).toBe(0);
    const sheet = sheet1(ods);
    expect(sheet.tables).toHaveLength(1);
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('Column 1'));
  });

  it('data-validation persists a list rule', () => {
    const ods = path2('dv.ods');
    run('new', ods);
    expect(run('data-validation', ods, 'A1:A5', '--list', 'Red,Green,Blue', '--in-place').exit).toBe(0);
    const rules = sheet1(ods).validations;
    expect(rules).toHaveLength(1);
    expect(rules[0]!.listValues).toEqual(['Red', 'Green', 'Blue']);
  });

  it('data-validation persists a range-source rule', () => {
    const ods = path2('dvr.ods');
    run('new', ods);
    expect(run('data-validation', ods, 'A1:A5', '--source', 'E1:E8', '--in-place').exit).toBe(0);
    const rules = sheet1(ods).validations;
    expect(rules).toHaveLength(1);
    expect(rules[0]!.listValues ?? null).toBeNull();
    expect(rules[0]!.listSource).toBeTruthy();
  });

  it('named ranges: add, list, and remove', () => {
    const ods = path2('names.ods');
    run('new', ods);
    expect(run('name', 'add', ods, 'TaxRate', 'B1:B10', '--in-place').exit).toBe(0);
    expect(run('name', 'list', ods).out).toContain('TaxRate');
    expect(run('name', 'remove', ods, 'TaxRate', '--in-place').exit).toBe(0);
    expect(run('name', 'list', ods).out).not.toContain('TaxRate');
  });

  it('hyperlink sets and removes a link without mangling the value', () => {
    const ods = path2('link.ods');
    run('new', ods);
    run('set', ods, 'A1', 'docs', '--in-place');
    expect(run('hyperlink', ods, 'A1', 'https://example.com', '--in-place').exit).toBe(0);
    expect(run('get', ods, 'A1').out.trim()).toBe('docs');
    expect(sheet1(ods).find(0, 0)!.hyperlink!.target).toBe('https://example.com');
    run('hyperlink', ods, 'A1', '--remove', '--in-place');
    expect(sheet1(ods).find(0, 0)!.hyperlink).toBeNull();
  });

  it('insert and delete rows shift data', () => {
    const ods = threeRowBook();
    run('insert-rows', ods, '2', '--in-place');
    expect(run('get', ods, 'A2').out.trim()).toBe('');
    expect(run('get', ods, 'A3').out.trim()).toBe('Beta');
    run('delete-rows', ods, '2', '--in-place');
    expect(run('get', ods, 'A2').out.trim()).toBe('Beta');
  });

  it('insert-cols accepts a letter', () => {
    const ods = threeRowBook();
    run('insert-cols', ods, 'A', '--in-place');
    expect(run('get', ods, 'B1').out.trim()).toBe('Name');
  });

  it('hide and show columns round-trip visibility', () => {
    const ods = threeRowBook();
    expect(run('hide-cols', ods, 'B', '--in-place').exit).toBe(0);
    expect(sheet1(ods).isColumnHidden(1)).toBe(true);
    expect(run('show-cols', ods, 'B', '--in-place').exit).toBe(0);
    expect(sheet1(ods).isColumnHidden(1)).toBe(false);
  });

  it('merge and unmerge a region', () => {
    const ods = path2('merge.ods');
    run('new', ods);
    run('set', ods, 'A1', 'wide', '--in-place');
    expect(run('merge', ods, 'A1:B2', '--in-place').exit).toBe(0);
    expect(sheet1(ods).findMerge(0, 0)).toBeDefined();
    expect(run('unmerge', ods, 'A1', '--in-place').exit).toBe(0);
    expect(sheet1(ods).findMerge(0, 0)).toBeUndefined();
  });

  it('data edits refuse to overwrite without --in-place or --output', () => {
    const ods = threeRowBook();
    const r = run('sort', ods, 'A1:B3', '--by', 'A');
    expect(r.exit).toBe(1);
    expect(r.err).toContain('--in-place');
  });
});

describe('chart sheets', () => {
  function chartDataBook(): string {
    const ods = path2('chartdata.ods');
    run('new', ods);
    run('set', ods, 'A1', 'Region', '--in-place');
    run('set', ods, 'B1', 'Q1', '--in-place');
    run('set', ods, 'C1', 'Q2', '--in-place');
    run('set', ods, 'A2', 'North', '--in-place');
    run('set', ods, 'B2', '10', '--in-place');
    run('set', ods, 'C2', '5', '--in-place');
    run('set', ods, 'A3', 'South', '--in-place');
    run('set', ods, 'B3', '20', '--in-place');
    run('set', ods, 'C3', '8', '--in-place');
    return ods;
  }

  it('a chart from a range seeds categories and a series per column', () => {
    const ods = chartDataBook();
    expect(
      run('chart', ods, '--type', 'column', '--title', 'Sales', '--from', 'Sheet1.A1:C3', '--in-place').exit,
    ).toBe(0);
    const sheet = sheet1(ods, 'Chart1');
    expect(sheet.isChartSheet).toBe(true);
    const chart = sheet.chart!;
    expect(chart.title).toBe('Sales');
    expect(chart.type).toBe('column');
    expect(chart.categoriesReference).toBe('Sheet1.A2:A3');
    expect(chart.series).toHaveLength(2);
    expect(chart.series[0]!.valuesReference).toBe('Sheet1.B2:B3');
    expect(chart.series[0]!.nameReference).toBe('Sheet1.B1');
  });

  it('charts can be listed, shown, updated, and removed', () => {
    const ods = chartDataBook();
    run('chart', ods, '--type', 'column', '--title', 'Sales', '--from', 'Sheet1.A1:C3', '--in-place');

    expect(run('charts', 'list', ods).out).toContain('Chart1\tcolumn\tSales\t2 series');
    let doc = json(run('charts', 'show', ods, 'Chart1').out);
    expect(doc.categories).toBe('Sheet1.A2:A3');
    expect(doc.series).toHaveLength(2);

    expect(
      run(
        'charts',
        'update',
        ods,
        'Chart1',
        '--type',
        'line',
        '--title',
        'Trend',
        '--legend',
        'false',
        '--in-place',
      ).exit,
    ).toBe(0);
    doc = json(run('charts', 'show', ods, 'Chart1').out);
    expect(doc.type).toBe('line');
    expect(doc.title).toBe('Trend');
    expect(doc.legend).toBe(false);
    expect(doc.series).toHaveLength(2);

    expect(
      run('charts', 'update', ods, 'Chart1', '--values', 'Sheet1.B2:B3', '--names', 'Only', '--in-place')
        .exit,
    ).toBe(0);
    doc = json(run('charts', 'show', ods, 'Chart1').out);
    expect(doc.series).toHaveLength(1);
    expect(doc.title).toBe('Trend');
    expect(doc.categories).toBe('Sheet1.A2:A3');

    expect(run('charts', 'update', ods, 'Chart1', '--legend', 'maybe', '--in-place').exit).toBe(1);
    expect(run('charts', 'show', ods, 'Sheet1').exit).toBe(1);
    expect(run('charts', 'remove', ods, 'Chart1', '--in-place').exit).toBe(0);
    expect(run('charts', 'list', ods, '--json').out.trim()).toBe('[]');
  });

  it('a chart with explicit fields round-trips', () => {
    const ods = chartDataBook();
    expect(
      run(
        'chart',
        ods,
        '--type',
        'bar',
        '--stacking',
        'percent',
        '--categories',
        'Sheet1.A2:A3',
        '--values',
        'Sheet1.B2:B3',
        '--names',
        'First',
        '--values',
        'Sheet1.C2:C3',
        '--names',
        'Second',
        '--y-title',
        'Sales',
        '--no-legend',
        '--in-place',
      ).exit,
    ).toBe(0);
    const chart = sheet1(ods, 'Chart1').chart!;
    expect(chart.type).toBe('bar');
    expect(chart.stacking).toBe('percent');
    expect(chart.showLegend).toBe(false);
    expect(chart.yAxisTitle).toBe('Sales');
    expect(chart.series).toHaveLength(2);
    expect(chart.series[0]!.name).toBe('First');
    expect(chart.series[1]!.valuesReference).toBe('Sheet1.C2:C3');
  });

  it('a scatter chart keeps its x values per series', () => {
    const ods = chartDataBook();
    expect(
      run(
        'chart',
        ods,
        '--type',
        'scatter',
        '--x-values',
        'Sheet1.B2:B3',
        '--values',
        'Sheet1.C2:C3',
        '--in-place',
      ).exit,
    ).toBe(0);
    const series = sheet1(ods, 'Chart1').chart!.series;
    expect(series).toHaveLength(1);
    expect(series[0]!.xValuesReference).toBe('Sheet1.B2:B3');
    expect(series[0]!.valuesReference).toBe('Sheet1.C2:C3');
  });

  it('a bubble chart requires bubble sizes', () => {
    const ods = chartDataBook();
    const r = run(
      'chart',
      ods,
      '--type',
      'bubble',
      '--x-values',
      'Sheet1.A2:A3',
      '--values',
      'Sheet1.B2:B3',
      '--in-place',
    );
    expect(r.exit).toBe(1);
    expect(r.err).toContain('--bubble-sizes');
  });

  it('a pie chart rejects more than one series', () => {
    const ods = chartDataBook();
    const r = run(
      'chart',
      ods,
      '--type',
      'pie',
      '--values',
      'Sheet1.B2:B3',
      '--values',
      'Sheet1.C2:C3',
      '--in-place',
    );
    expect(r.exit).toBe(1);
    expect(r.err).toContain('single series');
  });

  it('stacking is rejected for chart types that cannot stack', () => {
    const ods = chartDataBook();
    const r = run(
      'chart',
      ods,
      '--type',
      'line',
      '--from',
      'Sheet1.A1:C3',
      '--stacking',
      'stacked',
      '--in-place',
    );
    expect(r.exit).toBe(1);
    expect(r.err.toLowerCase()).toContain('stacking');
  });
});

describe('cellestra <file> — the launcher', () => {
  it('launching a workbook that does not exist is refused before anything launches', () => {
    const r = run(path2('nope.ods'));
    expect(r.exit).toBe(1);
    expect(r.err).toContain('does not exist');
    expect(launched).toHaveLength(0);
  });

  it('an argument that is not a workbook names both readings', () => {
    const r = run('recal');
    expect(r.exit).toBe(1);
    expect(r.err).toContain("isn't a workbook Cellestra opens");
    expect(r.err).toContain('cellestra --help');
  });

  it('bare cellestra still prints help rather than launching', () => {
    const r = run();
    expect(r.exit).toBe(0);
    expect(r.out).toContain('Usage');
    expect(r.out).toContain('recalc');
    expect(launched).toHaveLength(0);
  });

  it('an existing workbook goes to the launcher (stubbed here)', () => {
    const ods = path2('book.ods');
    run('new', ods);
    expect(run(ods).exit).toBe(0);
    expect(launched).toEqual([ods]);
  });
});

describe('review fixes', () => {
  it('--in-place keeps the mode, a symlink, and a hard link', () => {
    const real = path2('real.ods');
    const alias = path2('alias.ods');
    const twin = path2('twin.ods');
    run('new', real);
    chmodSync(real, 0o600);
    symlinkSync(real, alias);
    linkSync(real, twin);
    expect(run('set', alias, 'A1', '7', '--in-place').exit).toBe(0);
    expect(lstatSync(alias).isSymbolicLink()).toBe(true);
    expect(statSync(real).mode & 0o777).toBe(0o600);
    expect(statSync(real).nlink).toBe(2);
    expect(sheet1(twin).getValue(0, 0)).toEqual(CellValue.number(7));
    expect(readdirSync(dir).filter((n) => n.endsWith('.tmp'))).toEqual([]);
  });

  it('a new file written with -o gets the usual mode', () => {
    const out = path2('fresh.ods');
    run('new', out);
    expect(statSync(out).mode & 0o777).toBe(0o644 & ~process.umask());
  });

  it('handles ranges with more values than the stack holds arguments', () => {
    const ods = path2('big.ods');
    const book = Workbook.createNew();
    for (let r = 0; r < 200_000; r++) book.activeSheet.setValue(r, 0, CellValue.number(r - 5));
    writeFileSync(ods, writeOds(book));
    const r = run('stats', ods, 'A1:A200000', '--json');
    expect(r.exit).toBe(0);
    const stats = json(r.out);
    expect(stats.min).toBe(-5);
    expect(stats.max).toBe(199_994);
    expect(run('clear', ods, 'A1:A200000', '--in-place').exit).toBe(0);
    // Writing, reading and clearing 200,000 rows takes about 5 s on its own.
  }, 30_000);

  it('refuses inserts that would push data off the sheet or reach past it', () => {
    const ods = path2('edge.ods');
    run('new', ods);
    run('set', ods, 'A1=1', 'A1048576=2', '--in-place');
    const ops = (o: object) => {
      const file = path2('ops.json');
      writeFileSync(file, JSON.stringify([o]));
      return run('apply', ods, file, '--in-place');
    };
    const pushed = ops({ op: 'insert-rows', at: 1, count: 5 });
    expect(pushed.exit).toBe(1);
    expect(pushed.err).toContain('off the sheet');
    expect(sheet1(ods).getValue(1048575, 0)).toEqual(CellValue.number(2));
    const huge = ops({ op: 'insert-rows', at: 1, count: 100000000 });
    expect(huge.exit).toBe(1);
    expect(huge.err).toContain('past the sheet');
    expect(run('insert-cols', ods, 'B', '--count', '20000', '--in-place').exit).toBe(1);
    expect(run('insert-rows', ods, '2', '--count', '3', '--in-place').exit).toBe(1);
    run('clear', ods, 'A1048576', '--in-place');
    expect(run('insert-rows', ods, '1', '--count', '3', '--in-place').exit).toBe(0);
    expect(sheet1(ods).getValue(3, 0)).toEqual(CellValue.number(1));
  });

  it('reads --json the way the parser does', () => {
    const ods = path2('j.ods');
    run('new', ods);
    run('set', ods, 'A1', '3', '--in-place');
    expect(() => json(run('get', ods, 'A1', '--json=true').out)).not.toThrow();
    expect(run('get', ods, 'A1', '--json', 'false').out.trim()).toBe('3');
    expect(run('get', ods, 'A1', '--json=false').out.trim()).toBe('3');
    const bad = run('get', ods, 'A1', '--json=true', '--nope');
    expect(json(bad.out).ok).toBe(false);
  });
});

describe('rotation, conditional formats, page setup and print', () => {
  const book = () => {
    const ods = path2('book.ods');
    run('new', ods);
    writeFileSync(path2('d.csv'), 'Item,Q1\nA,5\nB,15\n');
    run('put', ods, 'A1', '--from', path2('d.csv'), '-i');
    return ods;
  };

  it('style --rotation turns text and 0 takes it off', () => {
    const ods = book();
    expect(run('style', ods, 'A1:B1', '--rotation', '450', '-i').exit).toBe(0);
    expect(json(run('get', ods, 'B1', '--json').out).style.rotation).toBe(90);
    run('style', ods, 'A1:B1', '--rotation', '0', '-i');
    expect(json(run('get', ods, 'B1', '--json').out)).not.toHaveProperty('style');
  });

  it('conditional adds, replaces and clears rules, and they survive a save', () => {
    const ods = book();
    expect(run('conditional', ods, 'B2:B3', '--when', '>', '--value', '10', '-i').exit).toBe(0);
    run(
      'conditional',
      ods,
      'B2:B3',
      '--when',
      'between',
      '--value',
      '1',
      '--value2',
      '6',
      '--format',
      'yellow',
      '-i',
    );
    expect(json(run('get', ods, 'B3', '--json').out).style.conditional).toEqual([
      { when: '>', value: '10', format: 'light-red' },
      { when: 'between', value: '1', value2: '6', format: 'yellow' },
    ]);
    run(
      'conditional',
      ods,
      'B2:B3',
      '--when',
      'formula',
      '--value',
      '=B2>B$2',
      '--format',
      'bold',
      '--replace',
      '-i',
    );
    const rules = json(run('get', ods, 'B3', '--json').out).style.conditional;
    expect(rules).toHaveLength(1);
    expect(rules[0].when).toBe('formula');
    expect(rules[0].value).toBe('=B2>B$2');
    expect(run('conditional', ods, 'B2:B3', '--clear', '-i').exit).toBe(0);
    expect(json(run('get', ods, 'B3', '--json').out)).not.toHaveProperty('style');
    expect(run('conditional', ods, 'B2', '--when', 'sometimes', '-i').exit).not.toBe(0);
  });

  it('page-setup patches the setup and sheets reports it', () => {
    const ods = book();
    run(
      'page-setup',
      ods,
      '--paper',
      'letter',
      '--orientation',
      'landscape',
      '--fit-width',
      '1',
      '--gridlines',
      '-i',
    );
    run('page-setup', ods, '--margins', '15', '--center', 'both', '-i');
    const setup = json(run('sheets', ods, '--json').out)[0].pageSetup;
    expect(setup).toMatchObject({
      paper: 'letter',
      orientation: 'landscape',
      margins: 15,
      fitWidth: 1,
      fitHeight: 0,
      gridlines: true,
      center: 'both',
    });
    expect(run('page-setup', ods, '--margins', '200', '-i').exit).not.toBe(0);
    run('page-setup', ods, '--reset', '-i');
    expect(json(run('sheets', ods, '--json').out)[0]).not.toHaveProperty('pageSetup');
  });

  it('apply takes conditional and page-setup ops', () => {
    const ods = book();
    const ops = path2('ops.json');
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'conditional', range: 'B2:B3', rules: [{ when: '<', value: 10, format: 'green' }] },
        { op: 'page-setup', paper: '200x300', pages: 2, headers: true },
        { op: 'style', range: 'A1', rotation: 45 },
      ]),
    );
    expect(run('apply', ods, ops, '-i').exit).toBe(0);
    expect(json(run('get', ods, 'B2', '--json').out).style.conditional[0].format).toBe('green');
    expect(json(run('get', ods, 'A1', '--json').out).style.rotation).toBe(45);
    expect(json(run('sheets', ods, '--json').out)[0].pageSetup).toMatchObject({
      paper: '200x300',
      pages: 2,
      headers: true,
    });
  });

  it('print writes printable HTML with the setup and the rules applied', () => {
    const ods = book();
    run('page-setup', ods, '--paper', 'a5', '--headers', '-i');
    run('conditional', ods, 'B2:B3', '--when', '>', '--value', '10', '-i');
    const r = run('print', ods, '-');
    expect(r.exit).toBe(0);
    expect(r.out).toContain('size:148mm 210mm');
    expect(r.out).toMatch(/>15</);
    expect(r.out.toLowerCase()).toContain('#ffc7ce');
    const out = path2('p.html');
    const j = json(run('print', ods, out, '--range', 'A1:A3', '--json').out);
    expect(j).toMatchObject({ output: out, pages: 1, format: 'html' });
    expect(readFileSync(out, 'utf8')).not.toMatch(/>15</);
    expect(run('print', ods, path2('p.txt')).exit).not.toBe(0);
  });
});

describe('isExpectedError', () => {
  it('treats user and file-system errors as expected, and runtime bugs as bugs', () => {
    const fsError = Object.assign(new Error('ENOENT: no such file'), { code: 'ENOENT' });
    const nodeError = Object.assign(new TypeError('bad arg'), { code: 'ERR_INVALID_ARG_TYPE' });
    const internal = Object.assign(new Error('internal'), { code: 'ERR_INTERNAL_ASSERTION' });
    expect(isExpectedError(fsError)).toBe(true);
    expect(isExpectedError(new RangeError("Sheet name can't be empty"))).toBe(true);
    expect(isExpectedError(new SyntaxError("'Q' is not a valid cell reference."))).toBe(true);
    expect(isExpectedError(new Error('duplicate sheet'))).toBe(true);
    expect(isExpectedError(nodeError)).toBe(false);
    expect(isExpectedError(internal)).toBe(false);
    expect(isExpectedError(new TypeError('x is undefined'))).toBe(false);
    expect(isExpectedError(new RangeError('Maximum call stack size exceeded'))).toBe(false);
    expect(isExpectedError(new RangeError('Invalid array length'))).toBe(false);
  });
});
