import { readFileSync, writeFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { SAVE_TRANSITIONS, step, toMermaid, type SaveEvent, type SaveState } from '../src/app/saveChart.ts';
import { runRandomOps, type Op } from './support/randomOps.ts';

/** The chart's state, plus a model of the file, the timers and the user's answers around it. */
interface World {
  state: SaveState;
  /** Writes started and not yet answered. */
  writes: number;
  /** Whether the write running is to the document's own file, the only kind that can meet a conflict. */
  inPlace: boolean;
  /** Writes that finished, so a read can tell one finished while it waited. */
  writesDone: number;
  /** Reads in progress, and how many writes had finished when each started. */
  reads: { valid: boolean; writesDone: number }[];
  /** Edits the file doesn't have yet. */
  unsaved: boolean;
  /** Edits made since the last write started. */
  editedSinceWrite: boolean;
  /** Auto-save's countdown is running. */
  timer: boolean;
  /** The conflict dialog is up and the user hasn't chosen. */
  conflictAsked: boolean;
  /** A close waits for the write in progress. */
  closeHeld: boolean;
  /** The user said to discard unsaved changes on close. */
  agreedToLose: boolean;
}

const busy = (state: SaveState) =>
  state === 'saving' || state === 'savingDirty' || state === 'copying' || state === 'conflict';

function fire(w: World, event: SaveEvent): void {
  const from = w.state;
  const { to, effects } = step(from, event);
  if (from === 'closed') expect(effects, 'nothing happens once closed').toEqual([]);
  if (from === 'conflict' && to !== 'conflict') w.conflictAsked = false;
  w.state = to;
  for (const effect of effects) {
    switch (effect) {
      case 'startTimer':
        w.timer = true;
        break;
      case 'writeFile':
      case 'pickFile':
        expect(w.writes, 'never two writes at once').toBe(0);
        expect(w.conflictAsked, 'no write while the conflict dialog waits for a choice').toBe(false);
        w.writes++;
        w.inPlace = effect === 'writeFile';
        w.editedSinceWrite = false;
        break;
      case 'showConflict':
        w.conflictAsked = true;
        break;
      case 'readFile':
        w.reads.push({ valid: true, writesDone: w.writesDone });
        break;
      case 'dropRead':
        for (const r of w.reads) r.valid = false;
        break;
      case 'takeFile': {
        const read = w.reads.find((r) => r.valid);
        expect(read, 'a file is taken only from a read that is still wanted').toBeDefined();
        expect(w.unsaved || w.writes > 0, 'a refresh never replaces unsaved data').toBe(false);
        expect(read?.writesDone, 'a refresh never takes a read that a save finished during').toBe(
          w.writesDone,
        );
        break;
      }
      case 'replaceWithFile':
        expect(
          w.reads.some((r) => r.valid),
          'the file replaces the document only from a read still wanted',
        ).toBe(true);
        // The user chose to drop this document's changes.
        w.unsaved = false;
        w.editedSinceWrite = false;
        break;
      case 'warnChanged':
        expect(w.unsaved, 'the warning is for a document with unsaved edits').toBe(true);
        break;
      case 'closeLater':
        w.closeHeld = true;
        break;
      case 'closeWindow':
        expect(w.unsaved && !w.agreedToLose, 'closing never loses edits the user did not agree to lose').toBe(
          false,
        );
        break;
      case 'askClose':
      case 'cancelClose':
      case 'retryLater':
        break;
    }
  }
}

/** The lifecycle sends a held close again as soon as no write is running. */
function closeWhenIdle(w: World): void {
  if (w.closeHeld && w.writes === 0 && !busy(w.state) && w.state !== 'reloading') {
    w.closeHeld = false;
    fire(w, 'close');
  }
}

/** A write ends. */
function writeEnds(w: World, event: 'saved' | 'failed' | 'versionChanged'): void {
  w.writes--;
  if (event === 'saved') {
    w.writesDone++;
    w.unsaved = w.editedSinceWrite;
  }
  fire(w, event);
  closeWhenIdle(w);
}

/** Saves take turns: a request waits while a write or the conflict question that follows one runs. */
const turnFree = (w: World) => w.writes === 0 && !busy(w.state) && w.state !== 'reloading';

const ops: Op<World>[] = [
  {
    name: 'edit',
    weight: 4,
    // The close and conflict questions are modal, and a closed window takes no typing.
    when: (w) => w.state !== 'askingClose' && w.state !== 'conflict' && w.state !== 'closed',
    run: (w) => {
      w.unsaved = true;
      w.editedSinceWrite = true;
      fire(w, 'edit');
    },
  },
  { name: 'save', when: turnFree, run: (w) => fire(w, 'save') },
  { name: 'saveAs', when: turnFree, run: (w) => fire(w, 'saveAs') },
  { name: 'revert', when: turnFree, run: (w) => fire(w, 'reload') },
  {
    name: 'timer',
    when: (w) => w.timer,
    run: (w) => {
      w.timer = false;
      fire(w, 'timer');
    },
  },
  { name: 'refresh', run: (w) => fire(w, 'refresh') },
  { name: 'close', when: (w) => w.state !== 'askingClose', run: (w) => fire(w, 'close') },
  {
    name: 'closeAnswer',
    when: (w) => w.state === 'askingClose',
    run: (w, r) => {
      const answer = r.pick(['save', 'saveAs', 'discard', 'cancel'] as const);
      if (answer === 'discard') w.agreedToLose = true;
      fire(w, answer);
      return answer;
    },
  },
  {
    name: 'conflictChoice',
    when: (w) => w.conflictAsked,
    run: (w, r) => {
      const choice = r.pick(['overwrite', 'saveAs', 'reload', 'cancel'] as const);
      fire(w, choice);
      closeWhenIdle(w);
      return choice;
    },
  },
  { name: 'saved', weight: 2, when: (w) => w.writes > 0, run: (w) => writeEnds(w, 'saved') },
  { name: 'failed', when: (w) => w.writes > 0, run: (w) => writeEnds(w, 'failed') },
  {
    name: 'versionChanged',
    when: (w) => w.writes > 0 && w.inPlace,
    run: (w) => writeEnds(w, 'versionChanged'),
  },
  {
    // A read finishes. The lifecycle ignores one it dropped; the others arrive as `readDone` or `readFailed`.
    name: 'readEnds',
    when: (w) => w.reads.length > 0,
    run: (w, r) => {
      const [read] = w.reads.splice(r.int(w.reads.length), 1);
      if (!read?.valid) return 'dropped';
      w.reads.push(read);
      const event = r.chance(0.8) ? 'readDone' : 'readFailed';
      fire(w, event);
      w.reads.splice(w.reads.indexOf(read), 1);
      closeWhenIdle(w);
      return event;
    },
  },
];

describe('the save chart', () => {
  it('random events never write twice at once or during a conflict, and never lose an edit unasked', async () => {
    await runRandomOps<World>({
      steps: 200,
      setup: () => ({
        state: 'clean',
        writes: 0,
        inPlace: false,
        writesDone: 0,
        reads: [],
        unsaved: false,
        editedSinceWrite: false,
        timer: false,
        conflictAsked: false,
        closeHeld: false,
        agreedToLose: false,
      }),
      ops,
      check: (w) => {
        if (w.state === 'clean' || w.state === 'refreshing')
          expect(w.unsaved, 'saved means the file has every edit').toBe(false);
        if (w.state === 'saving' || w.state === 'savingDirty' || w.state === 'copying')
          expect(w.writes, 'a write is running').toBe(1);
        if (w.state === 'closed')
          expect(!w.unsaved || w.agreedToLose, 'closed with nothing unsaved, or the user agreed').toBe(true);
      },
    });
  });

  it('lists each move once', () => {
    const keys = SAVE_TRANSITIONS.map((t) => `${t.from} ${t.on}`);
    expect(new Set(keys).size).toBe(keys.length);
  });

  it('is drawn in docs/architecture.md', () => {
    const file = new URL('../../../docs/architecture.md', import.meta.url);
    const doc = readFileSync(file, 'utf8');
    const block = /```mermaid\n(stateDiagram-v2\n[\s\S]*?)\n```/;
    const drawn = block.exec(doc)?.[1];
    expect(drawn, 'docs/architecture.md needs a ```mermaid stateDiagram-v2 block').toBeDefined();
    // `just diagram` redraws it from the table.
    if (process.env.WRITE_DIAGRAM === '1' && drawn !== toMermaid())
      writeFileSync(
        file,
        doc.replace(block, () => '```mermaid\n' + toMermaid() + '\n```'),
      );
    else expect(drawn, 'the diagram is out of date: run `just diagram`').toBe(toMermaid());
  });
});
