import { CellValue, OdsWriter, readOds, Workbook } from '@cellestra/core';
import { signal } from '@preact/signals';
import { describe, expect, it } from 'vitest';
import { ConflictError, WorkbookSession, type DocumentRef, type OpenedDocument } from '@cellestra/ui';
import { DocumentLifecycle } from '../src/app/documentLifecycle.ts';
import type { LifecycleHost } from '../src/app/lifecycleHost.ts';
import type { DocumentStores } from '../src/app/documentStores.ts';
import { runRandomOps, type Op, type Random } from './support/randomOps.ts';

function odsWith(text: string): Uint8Array {
  const wb = Workbook.createNew();
  wb.worksheets[0]!.setValue(0, 0, CellValue.text(text));
  return OdsWriter.write(wb, null);
}

function textOf(bytes: Uint8Array): string {
  const value = readOds(bytes).worksheets[0]!.getValue(0, 0);
  return value.kind === 'text' ? value.value : '';
}

interface StoredFile {
  version: number;
  /** The text in A1 at each version the file has had. */
  readonly history: Map<number, string>;
}

/** A store call waiting for the test to let it through. */
interface Parked {
  readonly kind: 'save' | 'saveAs' | 'reload';
  release(random: Random): void;
}

/** A question waiting for the user's answer. */
interface Asked {
  readonly kind: 'discard' | 'conflict';
  answer(random: Random): string;
}

/** Files the test's store holds, and every store call and question still waiting. */
class GatedWorld {
  readonly files = new Map<string, StoredFile>();
  readonly parked: Parked[] = [];
  readonly asked: Asked[] = [];
  readonly errors: string[] = [];
  readonly thrown: unknown[] = [];
  readonly session = new WorkbookSession();
  readonly lifecycle: DocumentLifecycle;
  readonly native = {
    watch: () => {},
    confirmClose: () => this.closedWindow(),
    cancelClose: () => {},
  };
  edits = 0;
  copies = 0;
  closed = false;
  /** The user answered Discard, and hasn't typed since. */
  agreedToLose = false;
  /** The newest version this window has shown, per file. */
  readonly shown = new Map<string, number>();

  constructor() {
    this.files.set('docs:a.ods', { version: 1, history: new Map([[1, 'start']]) });
    const stores = {
      storeFor: () => ({
        save: (ref: DocumentRef, bytes: Uint8Array, options: { force: boolean }) =>
          this.#park('save', (random) => this.#write(ref, bytes, options.force, random)),
        reload: (ref: DocumentRef) => this.#reload(ref),
      }),
      canSaveInPlace: () => true,
      primary: {
        saveAs: (name: string, bytes: Uint8Array) =>
          this.#park('saveAs', (random) => {
            if (random.chance(0.2)) return null;
            if (random.chance(0.1)) throw new Error('disk full');
            const id = `docs:copy${++this.copies}.ods`;
            this.files.set(id, { version: 1, history: new Map([[1, textOf(bytes)]]) });
            return { store: 'nas', name, id, version: '1' } satisfies DocumentRef;
          }),
      },
    } as unknown as DocumentStores;
    const host: LifecycleHost = {
      session: this.session,
      stores,
      autoSaveSeconds: signal(30),
      native: () => this.native,
      questions: {
        discard: () => this.#ask('discard', ['save', 'discard', 'cancel'] as const),
        csvLoss: async () => 'csv' as const,
        conflict: () => this.#ask('conflict', ['overwrite', 'saveAs', 'reload', 'cancel'] as const),
      },
      commitOpenEdit: () => true,
      isEditing: () => false,
      dialogOpen: () => this.asked.length > 0,
      info: () => {},
      error: (e) => this.errors.push(e instanceof Error ? e.message : String(e)),
      recovery: { clear: async () => {}, keepAside: async () => {} },
      remember: async () => {},
      replaced: () => {},
      reloadInPlace: (run) => run(),
    };
    this.lifecycle = new DocumentLifecycle(host);
    this.lifecycle.adopt({ ref: this.#ref('docs:a.ods'), bytes: odsWith('start') });
  }

  #ref(id: string): DocumentRef {
    return { store: 'nas', name: id.slice(5), id, version: String(this.files.get(id)!.version) };
  }

  #park<T>(kind: Parked['kind'], finish: (random: Random) => T): Promise<T> {
    return new Promise((resolve, reject) => {
      this.parked.push({
        kind,
        release: (random) => {
          try {
            resolve(finish(random));
          } catch (e) {
            reject(e);
          }
        },
      });
    });
  }

  #write(ref: DocumentRef, bytes: Uint8Array, force: boolean, random: Random): DocumentRef {
    if (random.chance(0.15)) throw new Error('disk full');
    const file = this.files.get(ref.id)!;
    if (!force && String(file.version) !== ref.version)
      throw new ConflictError('changed', 'changed elsewhere');
    file.history.set(++file.version, textOf(bytes));
    return this.#ref(ref.id);
  }

  /** A read answers with the file as it was when the read started about half the time, as a slow disk can. */
  #reload(ref: DocumentRef): Promise<OpenedDocument> {
    const file = this.files.get(ref.id)!;
    const startVersion = file.version;
    return this.#park('reload', (random) => {
      if (random.chance(0.1)) throw new Error('read failed');
      const version = random.chance(0.5) ? startVersion : file.version;
      return {
        ref: { ...ref, version: String(version) },
        bytes: odsWith(file.history.get(version)!),
      };
    });
  }

  #ask<T extends string>(kind: Asked['kind'], answers: readonly T[]): Promise<T> {
    return new Promise((resolve) => {
      this.asked.push({
        kind,
        answer: (random) => {
          const answer = random.pick(answers);
          if (kind === 'discard' && answer === 'discard') this.agreedToLose = true;
          resolve(answer);
          return answer;
        },
      });
    });
  }

  closedWindow(): void {
    expect(
      !this.session.modified.peek() || this.agreedToLose,
      'the window closes only with nothing unsaved, or after Discard',
    ).toBe(true);
    this.closed = true;
  }

  /** Run a lifecycle call without waiting for it; a throw is a failure. */
  track(p: Promise<unknown>): void {
    p.catch((e: unknown) => this.thrown.push(e));
  }

  text(): string {
    const value = this.session.activeSheet.peek().getValue(0, 0);
    return value.kind === 'text' ? value.value : '';
  }
}

const settle = async () => {
  for (let i = 0; i < 5; i++) await new Promise((r) => setImmediate(r));
};

const open = (w: GatedWorld) => !w.closed;

const ops: Op<GatedWorld>[] = [
  {
    name: 'edit',
    weight: 4,
    // Question dialogs are modal.
    when: (w) => open(w) && w.asked.length === 0,
    run: (w) => {
      const text = `e${++w.edits}`;
      w.session.activeSheet.peek().setValue(0, 0, CellValue.text(text));
      w.session.touch();
      w.agreedToLose = false;
      return text;
    },
  },
  { name: 'save', when: open, run: (w) => w.track(w.lifecycle.save()) },
  { name: 'saveAs', when: open, run: (w) => w.track(w.lifecycle.saveAs('ods')) },
  { name: 'autoSave', when: open, run: (w) => w.track(w.lifecycle.autoSave()) },
  { name: 'revert', when: open, run: (w) => w.track(w.lifecycle.reloadFromStore()) },
  { name: 'close', when: open, run: (w) => w.track(w.lifecycle.requestClose()) },
  {
    name: 'outsideWrite',
    when: (w) => open(w) && w.session.document.peek() !== null,
    run: (w) => {
      const id = w.session.document.peek()!.id;
      const file = w.files.get(id)!;
      file.history.set(++file.version, `outside${file.version}`);
      w.track(w.lifecycle.externalChange());
      return `${id} v${file.version}`;
    },
  },
  {
    // The window closed: open the file again, in a new document with its own save chart.
    name: 'reopen',
    when: (w) => w.closed && w.parked.length === 0 && w.asked.length === 0,
    run: (w) => {
      const id = w.shown.has('docs:a.ods') ? 'docs:a.ods' : [...w.files.keys()][0]!;
      const file = w.files.get(id)!;
      w.closed = false;
      w.agreedToLose = false;
      w.lifecycle.adopt({
        ref: { store: 'nas', name: id.slice(5), id, version: String(file.version) },
        bytes: odsWith(file.history.get(file.version)!),
      });
      w.shown.set(id, file.version);
      return id;
    },
  },
  {
    name: 'release',
    weight: 4,
    when: (w) => w.parked.length > 0,
    run: (w, r) => {
      const [call] = w.parked.splice(r.int(w.parked.length), 1);
      call!.release(r);
      return call!.kind;
    },
  },
  {
    name: 'answer',
    weight: 2,
    when: (w) => w.asked.length > 0,
    run: (w, r) => {
      // The newest question is the one on top.
      const asked = w.asked.pop()!;
      return `${asked.kind} ${asked.answer(r)}`;
    },
  },
];

describe('the document lifecycle', () => {
  it('random edits, saves, outside writes and answers never lose an edit unasked or go back a version', async () => {
    await runRandomOps<GatedWorld>({
      steps: 120,
      setup: () => new GatedWorld(),
      ops,
      check: async (w) => {
        await settle();
        expect(w.thrown, 'no lifecycle call throws').toEqual([]);
        const unexpected = w.errors.filter(
          (e) => !/disk full|read failed|^Auto-save|changed elsewhere/.test(e),
        );
        expect(unexpected, 'only the store’s own failures are shown').toEqual([]);
        const writes = w.parked.filter((p) => p.kind !== 'reload').length;
        expect(writes, 'at most one write at a time').toBeLessThanOrEqual(1);
        const ref = w.session.document.peek();
        if (!ref) return;
        const version = Number(ref.version);
        expect(version, 'the window never goes back to an older version').toBeGreaterThanOrEqual(
          w.shown.get(ref.id) ?? 0,
        );
        w.shown.set(ref.id, version);
        if (w.lifecycle.saveState === 'clean')
          expect(w.session.modified.peek(), 'the chart says saved only when the session is').toBe(false);
        if (!w.session.modified.peek())
          expect(w.text(), 'saved means the window shows the file at its version').toBe(
            w.files.get(ref.id)!.history.get(version),
          );
      },
    });
  }, 60_000);
});
