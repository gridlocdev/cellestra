/**
 * Random runs of the edits the app and the CLI make, checked after each step: the workbook saves and
 * reopens unchanged, and undoing every step gives back where the run started. `just fuzz` runs more
 * seeds; a failure prints FUZZ_SEED=<n> to replay it.
 */
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import type { MergedRegion } from '../src/index.ts';
import {
  CellRange,
  CellValue,
  ClipboardBlock,
  DeleteSheetCommand,
  EditCommands,
  EditHistory,
  EditRefusedError,
  FillCommand,
  MoveLinesCommand,
  PasteBlockCommand,
  RecalculationEngine,
  RenameSheetCommand,
  SheetDimensions,
  SheetName,
  SortFilterCommands,
  StructuralEditCommand,
  Workbook,
  readOds,
  writeOds,
  type EditCommand,
  type PasteMode,
  type Worksheet,
} from '../src/index.ts';
import { RANDOM_TIMEOUT, runRandomOps, type Op, type Random } from './support/randomOps.ts';

const LOCALE = 'en-US';
/** Edits stay in a small corner so they meet each other: merges, formulas and moves overlap. */
const ROWS = 8;
const COLUMNS = 6;
const FORMULAS = ['=A1+B2', '=SUM(A1:C4)', '=A1&"x"', '=IF(B1>2,A1,C3)', '=COUNT(A:A)', '=Sheet2.A1*2'];
const NUMBERS = ['0', '-1', '3.5', '1e9', '12%', '2026-03-31', 'TRUE'];
const PASTE_MODES: readonly PasteMode[] = ['all', 'values', 'formats', 'formulas'];
const COLORS = [0xffff0000, 0xff00aa00, 0xff3355ff];

interface State {
  workbook: Workbook;
  history: EditHistory;
  engine: RecalculationEngine;
  /** Where Undo can go back to: the start, or the last reload or sheet-list change. */
  start: string;
  /** How many edits the history holds, so redo after undo-all can count. */
  steps: number;
}

/** The workbook as its saved parts say it, without the zip's dates or the file's metadata. */
function snapshot(workbook: Workbook): string {
  const parts = unzipSync(writeOds(workbook), {
    filter: (file) =>
      file.name === 'content.xml' || file.name === 'styles.xml' || file.name === 'settings.xml',
  });
  return (
    Object.keys(parts)
      .sort()
      // An empty element is the same either way it is written. One tag per line, so a failure's diff
      // points at the element that changed.
      .map((name) => {
        const xml = strFromU8(parts[name]!).replace(/<([\w:-]+)><\/\1>/g, '<$1/>');
        return `${name}\n${xml.replaceAll('><', '>\n<')}`;
      })
      .join('\n')
  );
}

function adopt(s: State, workbook: Workbook): void {
  s.workbook = workbook;
  s.engine = new RecalculationEngine(workbook);
  s.engine.rebuildGraph();
  s.engine.recalculateAll();
  s.history = new EditHistory(workbook);
  s.start = snapshot(workbook);
  s.steps = 0;
}

const sheets = (s: State): readonly Worksheet[] => s.workbook.worksheets.filter((w) => !w.isChartSheet);
const sheet = (s: State, r: Random): Worksheet => r.pick(sheets(s));
const cell = (r: Random): [number, number] => [r.int(ROWS), r.int(COLUMNS)];
function range(r: Random, maxRows = 3, maxColumns = 3): CellRange {
  const top = r.int(ROWS);
  const left = r.int(COLUMNS);
  return CellRange.fromBounds(top, left, top + r.int(maxRows), left + r.int(maxColumns));
}

/** What a user might type: text that breaks naive code, numbers, dates or formulas. */
function input(r: Random): string {
  switch (r.int(3)) {
    case 0:
      return r.pick(FORMULAS);
    case 1:
      return r.pick(NUMBERS);
    default:
      return r.text();
  }
}

/** Run one command as the app does: through the history, then recalculated. A refused edit changes nothing. */
function run(s: State, command: EditCommand): string {
  try {
    s.engine.apply(s.history.execute(command));
  } catch (e) {
    if (e instanceof EditRefusedError) return `refused: ${e.message}`;
    throw e;
  }
  s.steps = Math.min(s.steps + 1, EditHistory.maxDepth);
  return command.label;
}

/** A name the rename dialog accepts: valid and not already taken. */
function sheetName(s: State, r: Random): string {
  const name = r.chance(0.5) ? r.text(12) : `S${r.int(100)}`;
  if (SheetName.validationError(name) || s.workbook.findSheet(name))
    return `N${s.workbook.sheetCount}x${r.int(1e6)}`;
  return name;
}

const ops: Op<State>[] = [
  {
    name: 'type',
    weight: 6,
    run: (s, r) => {
      const ws = sheet(s, r);
      const [row, column] = cell(r);
      const text = input(r);
      return `${run(s, EditCommands.setInput(ws.name, row, column, text, LOCALE, s.workbook))} ${JSON.stringify(text)}`;
    },
  },
  {
    name: 'setBlock',
    weight: 2,
    run: (s, r) => {
      const values = Array.from({ length: 1 + r.int(3) }, () =>
        Array.from({ length: 1 + r.int(3) }, () =>
          r.chance(0.5) ? CellValue.number(r.int(100) - 50) : CellValue.text(r.text()),
        ),
      );
      const [row, column] = cell(r);
      return run(s, EditCommands.setBlock(sheet(s, r).name, row, column, values));
    },
  },
  { name: 'clearContents', run: (s, r) => run(s, EditCommands.clearContents(sheet(s, r).name, range(r))) },
  { name: 'clearAll', weight: 0.5, run: (s, r) => run(s, EditCommands.clearAll(sheet(s, r).name, range(r))) },
  {
    name: 'style',
    weight: 2,
    run: (s, r) => {
      const ws = sheet(s, r);
      // A whole column or row styles the line, the other path through applyStyle.
      const target = r.chance(0.2)
        ? r.chance(0.5)
          ? CellRange.fromBounds(0, 1, SheetDimensions.maxRows - 1, 2)
          : CellRange.fromBounds(1, 0, 2, SheetDimensions.maxColumns - 1)
        : range(r);
      const bold = r.chance(0.5);
      const color = r.pick(COLORS);
      return run(
        s,
        EditCommands.applyStyle(s.workbook, ws.name, target, (style) =>
          bold
            ? { ...style, font: { ...style.font, bold: !style.font.bold } }
            : { ...style, fill: { background: color } },
        ),
      );
    },
  },
  {
    name: 'clearFormatting',
    weight: 0.5,
    run: (s, r) => run(s, EditCommands.clearFormatting(s.workbook, sheet(s, r).name, range(r))),
  },
  {
    name: 'resize',
    run: (s, r) =>
      run(
        s,
        r.chance(0.5)
          ? EditCommands.setColumnWidth(sheet(s, r).name, r.int(COLUMNS), 20 + r.int(200))
          : EditCommands.setRowHeight(sheet(s, r).name, r.int(ROWS), 10 + r.int(60)),
      ),
  },
  {
    name: 'hide',
    weight: 0.5,
    run: (s, r) => {
      const first = r.int(COLUMNS);
      const last = first + r.int(2);
      const hidden = r.chance(0.7);
      return run(
        s,
        r.chance(0.5)
          ? EditCommands.setColumnsHidden(sheet(s, r).name, first, last, hidden)
          : EditCommands.setRowsHidden(sheet(s, r).name, first, last, hidden),
      );
    },
  },
  {
    name: 'freeze',
    weight: 0.5,
    run: (s, r) => run(s, EditCommands.setFrozenPanes(sheet(s, r).name, r.int(3), r.int(3))),
  },
  {
    // The app offers Merge only when the selection overlaps no merge.
    name: 'merge',
    run: (s, r) => {
      const ws = sheet(s, r);
      const target = range(r);
      if (target.rowCount * target.columnCount < 2) return 'single cell';
      if (ws.merges.some((m) => m.range.intersects(target))) return 'overlaps a merge';
      return run(s, EditCommands.mergeCells(ws.name, target));
    },
  },
  {
    name: 'unmerge',
    when: (s) => sheets(s).some((w) => w.merges.length > 0),
    run: (s, r) => {
      const ws = r.pick(sheets(s).filter((w) => w.merges.length > 0));
      const merge: MergedRegion = r.pick(ws.merges);
      return run(s, EditCommands.unmergeCells(ws.name, merge.range.top, merge.range.left));
    },
  },
  {
    name: 'insertOrDelete',
    weight: 2,
    run: (s, r) =>
      run(
        s,
        new StructuralEditCommand(
          sheet(s, r).name,
          r.chance(0.5) ? 'row' : 'column',
          r.chance(0.5) ? 'insert' : 'delete',
          r.int(ROWS),
          1 + r.int(2),
        ),
      ),
  },
  {
    name: 'moveLines',
    run: (s, r) => {
      const first = 1 + r.int(4);
      return run(
        s,
        new MoveLinesCommand(
          sheet(s, r).name,
          r.chance(0.5) ? 'row' : 'column',
          first,
          first + r.int(2),
          r.chance(0.5) ? -1 : 1,
        ),
      );
    },
  },
  {
    // Copy or cut a block and paste it, as the app's clipboard and the CLI's copy and move do.
    name: 'copyPaste',
    weight: 2,
    run: (s, r) => {
      const from = sheet(s, r);
      const block = ClipboardBlock.capture(from, range(r), r.chance(0.3));
      const [row, column] = cell(r);
      return run(s, new PasteBlockCommand(sheet(s, r).name, [row, column], block, r.pick(PASTE_MODES)));
    },
  },
  {
    name: 'fill',
    run: (s, r) => {
      const source = range(r, 2, 2);
      const fill = CellRange.fromBounds(source.top, source.left, source.bottom + 1 + r.int(3), source.right);
      return run(s, new FillCommand(sheet(s, r).name, source, fill));
    },
  },
  {
    name: 'sort',
    run: (s, r) => {
      const target = range(r, 5, 3);
      return run(
        s,
        SortFilterCommands.sortByColumn(
          sheet(s, r).name,
          target,
          target.left + r.int(target.columnCount),
          r.chance(0.5),
          r.chance(0.5),
          LOCALE,
        ),
      );
    },
  },
  {
    name: 'renameSheet',
    weight: 0.5,
    run: (s, r) => {
      const ws = sheet(s, r);
      const name = sheetName(s, r);
      return `${run(s, new RenameSheetCommand(ws.name, name))} ${JSON.stringify(name)}`;
    },
  },
  {
    name: 'deleteSheet',
    weight: 0.3,
    when: (s) => sheets(s).length > 1,
    run: (s, r) => run(s, new DeleteSheetCommand(sheet(s, r).name)),
  },
  {
    // Adding a sheet clears the history in the app (PRD §8.5), so Undo stops there.
    name: 'addSheet',
    weight: 0.3,
    when: (s) => s.workbook.sheetCount < 4,
    run: (s) => {
      const added = s.workbook.addSheet();
      s.engine.recordStructuralChange();
      s.history.clear();
      s.start = snapshot(s.workbook);
      s.steps = 0;
      return added.name;
    },
  },
  {
    name: 'undo',
    weight: 2,
    run: (s) => {
      if (!s.history.canUndo) return 'nothing';
      s.engine.apply(s.history.undo());
      s.steps--;
    },
  },
  {
    name: 'redo',
    run: (s) => {
      if (!s.history.canRedo) return 'nothing';
      s.engine.apply(s.history.redo());
      s.steps++;
    },
  },
  {
    name: 'reload',
    weight: 0.2,
    run: (s) => adopt(s, readOds(writeOds(s.workbook))),
  },
  {
    name: 'undoEverything',
    weight: 0.3,
    run: (s) => {
      const before = snapshot(s.workbook);
      let undone = 0;
      while (s.history.canUndo) {
        s.engine.apply(s.history.undo());
        undone++;
      }
      // The history may have dropped its oldest steps; then the start can't be reached.
      if (undone === s.steps)
        expect(snapshot(s.workbook), 'undoing every step gives back the start').toBe(s.start);
      for (let i = 0; i < undone; i++) s.engine.apply(s.history.redo());
      expect(snapshot(s.workbook), 'redoing every step comes back').toBe(before);
      return `${undone} steps`;
    },
  },
];

describe('the engine', () => {
  it('random edits save and reopen unchanged, and undo back to the start', async () => {
    await runRandomOps<State>({
      steps: 60,
      setup: () => {
        const workbook = Workbook.createNew();
        workbook.addSheet('Sheet2');
        const s = { steps: 0 } as State;
        adopt(s, workbook);
        return s;
      },
      ops,
      check: (s) => {
        const saved = snapshot(s.workbook);
        expect(snapshot(readOds(writeOds(s.workbook))), 'the saved file opens back the same').toBe(saved);
      },
    });
  }, RANDOM_TIMEOUT);
});
