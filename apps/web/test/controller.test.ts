import { CellRange, CellValue, EditCommands, OdsWriter, SheetDimensions, Workbook } from '@cellestra/core';
import { effect } from '@preact/signals';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import {
  CELL_PAD_X,
  ConflictError,
  keepRecoveryAside,
  MAX_ROW,
  type DocumentRef,
  type FileStore,
  type GridController,
  type OpenedDocument,
} from '@cellestra/ui';
import type * as Ui from '@cellestra/ui';
import type { AppController as Controller, DialogState as DialogShown } from '../src/app/controller.ts';

// Recovery writes are no-ops here (no IndexedDB); watch which ones the controller asks for.
vi.mock('@cellestra/ui', async (importOriginal) => {
  const ui = await importOriginal<typeof Ui>();
  return { ...ui, keepRecoveryAside: vi.fn(ui.keepRecoveryAside) };
});

// The controller expects a browser; give it just enough of one (no IndexedDB: recovery is a no-op).
beforeAll(() => {
  const store = new Map<string, string>();
  vi.stubGlobal('localStorage', {
    getItem: (k: string) => store.get(k) ?? null,
    setItem: (k: string, v: string) => void store.set(k, v),
    removeItem: (k: string) => void store.delete(k),
  });
  vi.stubGlobal('matchMedia', () => ({ matches: false, addEventListener: () => {} }));
  vi.stubGlobal('document', { documentElement: { dataset: {} } });
});

async function makeApp(): Promise<Controller> {
  const { AppController } = await import('../src/app/controller.ts');
  return new AppController();
}

/** A NAS-like store whose saves resolve when the test says so. */
class FakeStore implements FileStore {
  readonly kind = 'nas' as const;
  readonly label = 'Fake';
  readonly capabilities = { saveInPlace: true, browse: true, watch: false };
  saves: { ref: DocumentRef; bytes: Uint8Array; finish: () => void }[] = [];
  version = 1;
  reloadBytes: Uint8Array = new Uint8Array();
  /** The version a reload finds on disk: by default, not the one opened. */
  reloadVersion: string | null = 'v-disk';
  reloadGate: Promise<void> = Promise.resolve();

  canSaveInPlace(): boolean {
    return true;
  }
  async reload(ref: DocumentRef): Promise<OpenedDocument> {
    await this.reloadGate;
    return { ref: { ...ref, version: this.reloadVersion }, bytes: this.reloadBytes };
  }
  save(ref: DocumentRef, bytes: Uint8Array): Promise<DocumentRef> {
    if (ref.version !== `v${this.version}`) return Promise.reject(new Error(`stale ${ref.version}`));
    return new Promise((resolve) => {
      this.saves.push({
        ref,
        bytes,
        finish: () => {
          this.version++;
          resolve({ ...ref, version: `v${this.version}` });
        },
      });
    });
  }
}

function withFake(app: Controller, fake: FakeStore): void {
  (app.stores as unknown as { nas: FileStore }).nas = fake;
  app.error = (e: unknown) => {
    throw e;
  };
}

const flush = () => new Promise((r) => setTimeout(r, 0));
const ref = (name: string): DocumentRef => ({ store: 'nas', name, id: `r:${name}`, version: 'v1' });
const csvBytes = new TextEncoder().encode('a,b\n1,2\n');

describe('blocked storage', () => {
  it('starts with defaults when every storage access throws', async () => {
    const real = globalThis.localStorage;
    const blocked = () => {
      throw new DOMException('denied', 'SecurityError');
    };
    vi.stubGlobal('localStorage', { getItem: blocked, setItem: blocked, removeItem: blocked });
    try {
      const app = await makeApp();
      expect(app.theme.value).toBe('system');
      app.theme.value = 'dark'; // writing it back fails quietly too
      expect(app.isDark.value).toBe(true);
    } finally {
      vi.stubGlobal('localStorage', real);
    }
  });
});

describe('saving', () => {
  it('keeps edits made during an in-flight save dirty', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    withFake(app, fake);
    app.session.load(OdsWriter.write(Workbook.createNew(), null), ref('a.ods'));
    app.session.touch();

    const saving = app.save();
    await flush();
    expect(fake.saves).toHaveLength(1);
    app.session.touch(); // an edit lands while the bytes are on their way
    fake.saves[0]!.finish();
    expect(await saving).toBe(true);
    expect(app.session.modified.peek()).toBe(true);
    expect(app.session.document.peek()?.version).toBe('v2');
  });

  it('runs a second save after the first instead of racing it', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    withFake(app, fake);
    app.session.load(OdsWriter.write(Workbook.createNew(), null), ref('a.ods'));
    app.session.touch();

    const first = app.save();
    await flush();
    app.session.touch();
    const second = app.save();
    await flush();
    expect(fake.saves).toHaveLength(1); // the second waits
    fake.saves[0]!.finish();
    await first;
    await flush();
    expect(fake.saves).toHaveLength(2);
    expect(fake.saves[1]!.ref.version).toBe('v2'); // no false "file changed" conflict
    fake.saves[1]!.finish();
    expect(await second).toBe(true);
    expect(app.session.modified.peek()).toBe(false);
  });

  it('still runs Save As or an export asked for while a save is in flight', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    withFake(app, fake);
    const picker = vi.fn(() => Promise.resolve(null));
    (app as unknown as { browser: { saveAs: typeof picker } }).browser.saveAs = picker;
    app.session.load(OdsWriter.write(Workbook.createNew(), null), ref('a.ods'));
    app.session.touch();

    const saving = app.save();
    await flush();
    const exporting = app.saveAs('csv');
    fake.saves[0]!.finish();
    await saving;
    await exporting;
    expect(picker).toHaveBeenCalledOnce();
  });

  it('a save that finishes after a new document leaves the new document untitled', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    withFake(app, fake);
    app.session.load(OdsWriter.write(Workbook.createNew(), null), ref('a.ods'));
    app.session.touch();

    const saving = app.save();
    await flush();
    app.session.newDocument();
    app.session.touch();
    fake.saves[0]!.finish();
    expect(await saving).toBe(true);
    expect(app.session.document.peek()).toBeNull();
    expect(app.session.modified.peek()).toBe(true);
  });

  it('answers an open question with its cancel value when another is asked', async () => {
    const app = await makeApp();
    const first = app.ask<'save' | 'discard' | 'cancel'>('cancel', (resolve) => ({
      kind: 'confirm-discard',
      resolve,
    }));
    expect(app.asking).toBe(true);
    const second = app.prompt('Rename', 'Name', 'x');
    expect(await first).toBe('cancel');
    expect(app.dialog.peek()?.kind).toBe('prompt');
    // Another dialog shown over it cancels it too.
    app.dialog.value = { kind: 'about' };
    expect(await second).toBeNull();
    expect(app.asking).toBe(false);
    expect(app.dialog.peek()?.kind).toBe('about');
  });

  it('makes a repeated save a no-op once the first one leaves the document clean', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    withFake(app, fake);
    app.session.load(OdsWriter.write(Workbook.createNew(), null), ref('a.ods'));
    app.session.touch();
    const first = app.save();
    const second = app.save();
    await flush();
    fake.saves[0]!.finish();
    expect(await first).toBe(true);
    expect(await second).toBe(true);
    expect(fake.saves).toHaveLength(1);
  });

  it('asks before a lossy CSV save, and can switch to .ods', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    withFake(app, fake);
    app.session.load(csvBytes, ref('data.csv'));
    app.session.addSheet();
    app.session.setActiveSheet(0); // CSV writes the active sheet

    const saving = app.save();
    await flush();
    const d = app.dialog.peek();
    expect(d?.kind).toBe('csv-loss');
    if (d?.kind !== 'csv-loss') return;
    expect(d.losses.join()).toContain('sheet');
    d.resolve('cancel');
    expect(await saving).toBe(false);
    expect(fake.saves).toHaveLength(0);

    const again = app.save();
    await flush();
    const d2 = app.dialog.peek();
    if (d2?.kind !== 'csv-loss') throw new Error('expected the CSV prompt');
    d2.resolve('csv');
    await flush();
    expect(fake.saves).toHaveLength(1);
    expect(new TextDecoder().decode(fake.saves[0]!.bytes)).toContain('a,b');
    fake.saves[0]!.finish();
    expect(await again).toBe(true);
  });

  it('saves a plain CSV without asking', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    withFake(app, fake);
    app.session.load(csvBytes, ref('data.csv'));
    app.session.execute(
      EditCommands.setValue(app.session.activeSheet.peek().name, 2, 0, CellValue.number(3)),
    );
    const saving = app.save();
    await flush();
    expect(app.dialog.peek()).toBeNull();
    fake.saves[0]!.finish();
    expect(await saving).toBe(true);
  });

  it('sends a browser file without a handle through Save As instead of a silent download', async () => {
    const app = await makeApp();
    app.session.load(csvBytes, { store: 'browser', name: 'up.csv', id: 'up.csv', version: '1' });
    app.session.touch();
    const browserSave = vi.spyOn(app.browser, 'save');
    const browserSaveAs = vi.spyOn(app.browser, 'saveAs').mockResolvedValue(null);
    expect(await app.save()).toBe(false);
    expect(browserSave).not.toHaveBeenCalled();
    expect(browserSaveAs).toHaveBeenCalledOnce();
  });

  it('keeps a document downloaded as a copy marked unsaved', async () => {
    const app = await makeApp();
    app.session.load(csvBytes, { store: 'browser', name: 'up.csv', id: 'up.csv', version: '1' });
    app.session.touch();
    vi.spyOn(app.browser, 'saveAs').mockResolvedValue({
      store: 'browser',
      name: 'up.csv',
      id: 'up.csv',
      version: null,
      downloaded: true,
    });
    vi.spyOn(app, 'askCsvLoss').mockResolvedValue('csv');
    expect(await app.save()).toBe(true);
    expect(app.session.modified.peek()).toBe(true);
    expect(app.toast.peek()?.text).toMatch(/^Downloaded up\.csv/);
  });
});

describe('overwriting after a conflict', () => {
  it('still meets the LibreOffice lock, and overrides it only when that dialog says so', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    withFake(app, fake);
    app.session.load(OdsWriter.write(Workbook.createNew(), null), ref('a.ods'));
    app.session.touch();
    const save = vi
      .spyOn(fake, 'save')
      .mockRejectedValueOnce(new ConflictError('changed', 'The file changed on the server'))
      .mockRejectedValueOnce(new ConflictError('locked', 'The file is open in LibreOffice'))
      .mockResolvedValueOnce({ ...ref('a.ods'), version: 'v9' });
    const answer = async (reason: string) => {
      await flush();
      const d = app.dialog.peek() as Extract<DialogShown, { kind: 'conflict' }>;
      expect(d.reason).toBe(reason);
      d.resolve('overwrite');
    };
    const saving = app.save();
    await answer('changed');
    await answer('locked');
    expect(await saving).toBe(true);
    expect(save.mock.calls.map((c) => (c as unknown[])[2])).toEqual([
      { force: false, ignoreLock: false },
      { force: true, ignoreLock: false },
      { force: true, ignoreLock: true },
    ]);
  });
});

describe('saving from the unsaved-changes question in a browser that can only download', () => {
  function downloaded(app: Controller, done: 'downloaded' | 'saved') {
    app.session.load(csvBytes, { store: 'browser', name: 'up.csv', id: 'up.csv', version: '1' });
    app.session.touch();
    vi.spyOn(app, 'askCsvLoss').mockResolvedValue('csv');
    vi.spyOn(app.browser, 'saveAs').mockResolvedValue({
      store: 'browser',
      name: 'up.csv',
      id: 'up.csv',
      version: null,
      ...(done === 'downloaded' ? { downloaded: true } : {}),
    });
  }
  const answerSave = (app: Controller) =>
    (app.dialog.peek() as Extract<DialogShown, { kind: 'confirm-discard' }>).resolve('save');

  it('keeps a recovery snapshot of the document it replaces, since the download is not confirmed', async () => {
    const app = await makeApp();
    downloaded(app, 'downloaded');
    vi.mocked(keepRecoveryAside).mockClear();
    const replacing = app.newDocument();
    answerSave(app);
    await replacing;
    expect(app.session.document.peek()).toBeNull();
    expect(keepRecoveryAside).toHaveBeenCalledOnce();
    expect(vi.mocked(keepRecoveryAside).mock.calls[0]![0]).toBe('up.csv');
  });

  it('keeps none once the file was really saved', async () => {
    const app = await makeApp();
    downloaded(app, 'saved');
    vi.mocked(keepRecoveryAside).mockClear();
    const replacing = app.newDocument();
    answerSave(app);
    await replacing;
    expect(keepRecoveryAside).not.toHaveBeenCalled();
  });
});

describe('recovery and reloads', () => {
  it('restores a CSV document so Save still writes CSV, not a ZIP', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    withFake(app, fake);
    const wb = Workbook.createNew();
    wb.activeSheet.setValue(0, 0, CellValue.text('x'));
    // What the snapshot holds for a CSV document: .ods bytes, the .csv ref.
    app.session.load(OdsWriter.write(wb, null), ref('data.csv'), 'data.ods', 'csv');
    app.session.touch();
    expect(app.session.format.peek()).toBe('csv');
    const saving = app.save();
    await flush();
    const bytes = fake.saves[0]!.bytes;
    expect(bytes[0]).not.toBe(0x50); // not "PK", a ZIP
    expect(new TextDecoder().decode(bytes)).toBe('x\n');
    fake.saves[0]!.finish();
    await saving;
  });

  it('drops an external-change reload when edits land while the file is read', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    withFake(app, fake);
    app.session.load(csvBytes, ref('data.csv'));
    let open = () => {};
    fake.reloadGate = new Promise((r) => (open = r));
    fake.reloadBytes = new TextEncoder().encode('changed\n');
    app.info = () => {};
    const reloading = app.externalChange();
    await flush();
    app.session.touch(); // the user types while the read is in flight
    open();
    await reloading;
    expect(app.session.modified.peek()).toBe(true);
    expect(app.session.activeSheet.peek().getValue(0, 0)).toEqual(CellValue.text('a'));
  });
});

describe('a file changed on disk (desktop)', () => {
  function opened(app: Controller): FakeStore {
    const fake = new FakeStore();
    withFake(app, fake);
    app.session.load(csvBytes, ref('data.csv'));
    fake.reloadBytes = new TextEncoder().encode('changed\n');
    return fake;
  }

  it('keeps the document and its undo history when the contents did not change', async () => {
    const app = await makeApp();
    const fake = opened(app);
    fake.reloadVersion = 'v1'; // touched or rewritten by a sync client, same bytes
    const reload = vi.spyOn(app.session, 'reload');
    await app.externalChange();
    expect(reload).not.toHaveBeenCalled();
    expect(app.session.activeSheet.peek().getValue(0, 0)).toEqual(CellValue.text('a'));
  });

  it('reloads a clean document whose contents changed', async () => {
    const app = await makeApp();
    opened(app);
    await app.externalChange();
    expect(app.session.activeSheet.peek().getValue(0, 0)).toEqual(CellValue.text('changed'));
  });

  it('waits while a dialog or a cell edit is open, then reloads once', async () => {
    vi.useFakeTimers();
    try {
      const app = await makeApp();
      opened(app);
      const reload = vi.spyOn(app, 'reloadFromStore');
      app.dialog.value = { kind: 'about' };
      await app.externalChange();
      await app.externalChange(); // a second event folds into the first
      await vi.advanceTimersByTimeAsync(5_000);
      expect(reload).not.toHaveBeenCalled();

      let editing = true;
      app.grid.current = { isEditing: () => editing, commitEdit: () => true } as unknown as GridController;
      app.dialog.value = null;
      await vi.advanceTimersByTimeAsync(5_000);
      expect(reload).not.toHaveBeenCalled();

      editing = false;
      await vi.advanceTimersByTimeAsync(5_000);
      expect(reload).toHaveBeenCalledOnce();
      expect(app.session.activeSheet.peek().getValue(0, 0)).toEqual(CellValue.text('changed'));
    } finally {
      vi.useRealTimers();
    }
  });
});

describe('grid focus at launch', () => {
  it('focuses the grid when focus sits on the page, not when a field or dialog has it', async () => {
    const app = await makeApp();
    const body = {};
    const doc = globalThis.document as unknown as { body: unknown; activeElement: unknown };
    doc.body = body;
    const focus = vi.fn();
    app.grid.current = {
      focus,
      beginEdit: () => {},
      isEditing: () => false,
      commitEdit: () => true,
      revealActive: () => {},
    };

    doc.activeElement = body;
    app.focusGridIfIdle();
    expect(focus).toHaveBeenCalledTimes(1);

    doc.activeElement = { tagName: 'INPUT' };
    app.focusGridIfIdle();
    expect(focus).toHaveBeenCalledTimes(1);

    doc.activeElement = body;
    app.dialog.value = { kind: 'recovery', name: 'x', resolve: () => {} } as never;
    app.focusGridIfIdle();
    expect(focus).toHaveBeenCalledTimes(1);
  });
});

describe('auto-save', () => {
  it('clamps the interval and treats 0 or less as off', async () => {
    const { normalizeAutoSave } = await import('../src/app/controller.ts');
    expect(normalizeAutoSave(0)).toBe(0);
    expect(normalizeAutoSave(-3)).toBe(0);
    expect(normalizeAutoSave(2)).toBe(5);
    expect(normalizeAutoSave(30)).toBe(30);
    expect(normalizeAutoSave(9000)).toBe(600);
    expect(normalizeAutoSave(Number.NaN)).toBe(0);
  });

  it('keeps the interval when the Settings field is cleared, and only a typed 0 turns it off', async () => {
    const { autoSaveFromInput } = await import('../src/app/controller.ts');
    expect(autoSaveFromInput('', 30)).toBe(30);
    expect(autoSaveFromInput('  ', 45)).toBe(45);
    expect(autoSaveFromInput('abc', 30)).toBe(30);
    expect(autoSaveFromInput('0', 30)).toBe(0);
    expect(autoSaveFromInput('2', 30)).toBe(5);
  });

  it('saves in place once editing pauses, and the countdown restarts on each edit', async () => {
    vi.useFakeTimers();
    try {
      const app = await makeApp();
      const fake = new FakeStore();
      withFake(app, fake);
      app.autoSaveSeconds.value = 10;
      app.session.load(OdsWriter.write(Workbook.createNew(), null), ref('a.ods'));
      app.session.touch();
      await vi.advanceTimersByTimeAsync(8_000);
      app.session.touch(); // still typing: push the save back
      await vi.advanceTimersByTimeAsync(8_000);
      expect(fake.saves).toHaveLength(0);
      await vi.advanceTimersByTimeAsync(3_000);
      expect(fake.saves).toHaveLength(1);
      fake.saves[0]!.finish();
      await vi.advanceTimersByTimeAsync(0);
      expect(app.session.modified.peek()).toBe(false);
      // Clean now: nothing re-arms.
      await vi.advanceTimersByTimeAsync(60_000);
      expect(fake.saves).toHaveLength(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it('never asks: a new document waits for ⌘S, and off means off', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    withFake(app, fake);
    app.session.load(OdsWriter.write(Workbook.createNew(), null), null);
    app.session.touch();
    expect(await app.autoSave()).toBe(false);
    expect(app.dialog.peek()).toBeNull();

    app.session.load(OdsWriter.write(Workbook.createNew(), null), ref('b.ods'));
    app.session.touch();
    app.autoSaveSeconds.value = 0;
    expect(await app.autoSave()).toBe(false);
    expect(fake.saves).toHaveLength(0);
  });

  it('stands down after a failed write until a save succeeds', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    const errors: unknown[] = [];
    withFake(app, fake);
    app.error = (e) => void errors.push(e);
    app.autoSaveSeconds.value = 30;
    app.session.load(OdsWriter.write(Workbook.createNew(), null), { ...ref('c.ods'), version: 'v9' });
    app.session.touch();
    expect(await app.autoSave()).toBe(false); // stale version → the fake store throws
    expect(errors).toHaveLength(1);
    expect(await app.autoSave()).toBe(false);
    expect(errors).toHaveLength(1); // reported once, not on every tick
  });

  it('says once when a conflict or a refused permission stops it, and stays down until ⌘S', async () => {
    for (const failure of [
      new ConflictError('changed', 'changed elsewhere'),
      new DOMException('no', 'NotAllowedError'),
    ]) {
      const app = await makeApp();
      const fake = new FakeStore();
      const errors: unknown[] = [];
      withFake(app, fake);
      app.error = (e) => void errors.push(e);
      let attempts = 0;
      fake.save = () => {
        attempts++;
        return Promise.reject(failure);
      };
      app.autoSaveSeconds.value = 30;
      app.session.load(OdsWriter.write(Workbook.createNew(), null), ref('d.ods'));
      app.session.touch();
      expect(await app.autoSave()).toBe(false);
      app.session.touch();
      expect(await app.autoSave()).toBe(false);
      expect(attempts).toBe(1); // no retry behind the user's back
      expect(errors).toHaveLength(1);
      expect(String(errors[0])).toMatch(/Auto-save .*paused until you save/);

      // Another document starts afresh.
      app.session.load(OdsWriter.write(Workbook.createNew(), null), ref('e.ods'));
      app.session.touch();
      await app.autoSave();
      expect(attempts).toBe(2);
    }
  });
});

describe('createTable', () => {
  async function tableFrom(rows: (string | number)[][], top = 0) {
    const app = await makeApp();
    const sheet = app.session.activeSheet.peek();
    rows.forEach((r, i) =>
      r.forEach((v, c) =>
        sheet.setValue(top + i, c, typeof v === 'number' ? CellValue.number(v) : CellValue.text(v)),
      ),
    );
    app.selection.select(top, 0);
    app.createTable();
    return { app, sheet, table: sheet.tables[0]! };
  }

  it('uses a first row of text labels as the header, even over text data', async () => {
    const { table } = await tableFrom([
      ['Name', 'City'],
      ['Ann', 'Oslo'],
    ]);
    expect(table.hasHeader).toBe(true);
    expect(table.range.toString()).toBe('A1:B2');
  });

  it('puts labels in a new row when the data starts with numbers', async () => {
    const { app, sheet, table } = await tableFrom([
      [1, 2],
      [3, 4],
    ]);
    expect(table.hasHeader).toBe(true);
    expect(table.range.toString()).toBe('A1:B3');
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('Column 1'));
    expect(sheet.getValue(1, 0)).toEqual(CellValue.number(1));
    app.session.undo();
    expect(sheet.tables).toHaveLength(0);
    expect(sheet.getValue(0, 0)).toEqual(CellValue.number(1));
  });

  it('uses the blank row above the data for the labels', async () => {
    const { sheet, table } = await tableFrom([[1, 2]], 3);
    expect(table.range.toString()).toBe('A3:B4');
    expect(sheet.getValue(2, 1)).toEqual(CellValue.text('Column 2'));
    expect(sheet.getValue(3, 0)).toEqual(CellValue.number(1));
  });
});

describe('find and replace one at a time', () => {
  async function appWith(cells: [number, number, string][]) {
    const app = await makeApp();
    const sheet = app.session.activeSheet.peek();
    for (const [r, c, t] of cells) sheet.setValue(r, c, CellValue.text(t));
    return { app, sheet };
  }

  it('lists every match with its address and what it shows, in row-major order', async () => {
    const { app, sheet } = await appWith([
      [3, 0, 'apple pie'],
      [0, 2, 'Apple'],
      [1, 1, 'pear'],
    ]);
    sheet.setFormula(5, 1, 'LEN("apple")', CellValue.number(5));
    expect(app.findAll('apple').map((m) => [m.address, m.text])).toEqual([
      ['C1', 'Apple'],
      ['A4', 'apple pie'],
      ['B6', '=LEN("apple")'],
    ]);
    expect(app.findAll('')).toEqual([]);
  });

  it('first goes to a match, then replaces it and moves on', async () => {
    const { app, sheet } = await appWith([
      [0, 1, 'cat'],
      [2, 0, 'a cat and a cat'],
      [4, 4, 'dog'],
    ]);
    // A1 holds no match: the first press only finds B1.
    expect(app.replaceNext('cat', 'dog')).toEqual({ replaced: false, found: true });
    expect(app.selection.active.peek()).toEqual({ row: 0, column: 1 });
    expect(app.replaceNext('cat', 'dog')).toEqual({ replaced: true, found: true });
    expect(sheet.getValue(0, 1)).toEqual(CellValue.text('dog'));
    expect(app.selection.active.peek()).toEqual({ row: 2, column: 0 });
    expect(app.replaceNext('cat', 'dog')).toEqual({ replaced: true, found: false });
    expect(sheet.getValue(2, 0)).toEqual(CellValue.text('a dog and a dog'));
    app.session.undo();
    expect(sheet.getValue(2, 0)).toEqual(CellValue.text('a cat and a cat'));
  });

  it('wraps around to the first match after the last', async () => {
    const { app } = await appWith([
      [1, 0, 'x'],
      [5, 0, 'x'],
    ]);
    app.selection.select(5, 0);
    expect(app.findNext('x')).toBe(true);
    expect(app.selection.active.peek()).toEqual({ row: 1, column: 0 });
    expect(app.findNext('x', true)).toBe(true);
    expect(app.selection.active.peek()).toEqual({ row: 5, column: 0 });
  });
});

describe('replace all', () => {
  it('keeps text cells text, and formulas formulas', async () => {
    const app = await makeApp();
    const sheet = app.session.activeSheet.peek();
    sheet.setValue(0, 0, CellValue.text('007'));
    sheet.setValue(1, 0, CellValue.text('x+1'));
    sheet.setValue(2, 0, CellValue.text('1/x'));
    sheet.setValue(4, 0, CellValue.text('1/2'));
    sheet.setFormula(3, 0, 'SUM(1;2)', CellValue.empty);
    expect(app.replaceAll('x', '=')).toBe(2);
    expect(app.replaceAll('7', '9')).toBe(1);
    expect(app.replaceAll('2', '3')).toBe(2);
    expect(app.replaceAll('SUM', 'MAX')).toBe(1);
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('009'));
    expect(sheet.getValue(1, 0)).toEqual(CellValue.text('=+1'));
    expect(sheet.find(1, 0)?.formula).toBeNull();
    expect(sheet.getValue(2, 0)).toEqual(CellValue.text('1/='));
    expect(sheet.find(3, 0)?.formula).toBe('MAX(1;3)');
    expect(sheet.getValue(4, 0)).toEqual(CellValue.text('1/3')); // not a date
  });
});

/**
 * Answer every yes-or-no question the app asks in its own dialog with `answer()`, recording what it
 * asked. Questions are answered as they open.
 */
function answerConfirms(app: Controller) {
  const asked: string[] = [];
  let answer = false;
  const stop = effect(() => {
    const d = app.dialog.value;
    if (d?.kind !== 'confirm') return;
    asked.push(`${d.title} ${d.message}`);
    queueMicrotask(() => d.resolve(answer));
  });
  return { asked, answer: (v: boolean) => (answer = v), stop };
}

describe('large removals', () => {
  async function filled(rows: number) {
    const app = await makeApp();
    const sheet = app.session.activeSheet.peek();
    for (let r = 0; r < rows; r++) sheet.setValue(r, 0, CellValue.number(r));
    return { app, sheet, ...answerConfirms(app) };
  }

  it('clears 100 filled cells without asking', async () => {
    const { app, sheet, asked } = await filled(100);
    app.selection.selectRange({ top: 0, left: 0, bottom: 999, right: 5 });
    await app.clearContents();
    expect(asked).toHaveLength(0);
    expect(sheet.getValue(5, 0)).toEqual(CellValue.empty);
  });

  it('asks in its own dialog before clearing or deleting more than 100 filled cells', async () => {
    const { app, sheet, asked, answer } = await filled(101);
    app.selection.selectRange({ top: 0, left: 0, bottom: 200, right: 0 });
    await app.clearContents();
    await app.clearAll();
    await app.deleteRows();
    await app.deleteColumns();
    expect(asked).toHaveLength(4);
    expect(asked[0]).toMatch(/101 cells/);
    expect(sheet.getValue(5, 0)).toEqual(CellValue.number(5));

    answer(true);
    app.selection.selectRange({ top: 0, left: 1, bottom: 0, right: 1 }); // B1: its column is empty
    await app.deleteColumns();
    expect(asked).toHaveLength(4);
    app.selection.selectRange({ top: 0, left: 0, bottom: 0, right: 0 });
    await app.deleteColumns();
    expect(asked).toHaveLength(5);
    expect(sheet.getValue(5, 0)).toEqual(CellValue.empty);
  });

  it('asks before resetting a sheet with more than 100 cells', async () => {
    const { app, sheet, asked, answer } = await filled(101);
    await app.resetSheetData();
    expect(asked).toHaveLength(1);
    expect(sheet.getValue(5, 0)).toEqual(CellValue.number(5));
    answer(true);
    await app.resetSheetData();
    expect(sheet.getValue(5, 0)).toEqual(CellValue.empty);
  });
});

describe('name box', () => {
  it('jumps to a named range, on its own sheet', async () => {
    const app = await makeApp();
    app.session.addSheet();
    const wb = app.session.workbook.peek();
    const other = wb.worksheets[wb.worksheets.length - 1]!.name;
    wb.addNamedRange({ name: 'Totals', target: CellRange.fromBounds(1, 1, 3, 2, other) });
    app.session.setActiveSheet(0);
    expect(app.goTo('totals')).toBe(true);
    expect(app.sheetName).toBe(other);
    expect(app.range.toString()).toBe('B2:C4');
    expect(app.goTo('Nope')).toBe(false);
  });
});

describe('chart sheets', () => {
  it('leave styles and CSV export alone', async () => {
    const { buildCommands } = await import('../src/app/commands.ts');
    const { createChart } = await import('@cellestra/charts');
    const app = await makeApp();
    const catalog = buildCommands(app);
    const exportCsv = catalog.get('file.exportCsv')!;
    expect(catalog.isEnabled(exportCsv)).toBe(true);
    app.insertChartSheet(createChart());
    expect(app.sheet.isChartSheet).toBe(true);
    expect(catalog.isEnabled(exportCsv)).toBe(false);
    const undoable = app.session.canUndo.peek();
    app.style('Bold', (st) => ({ ...st, font: { ...st.font, bold: true } }));
    expect(app.session.canUndo.peek()).toBe(undoable);
    expect(app.sheet.find(0, 0)).toBeUndefined();
  });
});

describe('zoom', () => {
  it('zooms by a pinch or ⌘-wheel factor, within the zoom range', async () => {
    const { wheelZoomFactor } = await import('../src/app/controller.ts');
    const app = await makeApp();
    app.zoomBy(wheelZoomFactor(-100, 0)); // pinch out / wheel up: bigger
    expect(app.zoom.value).toBeGreaterThan(1);
    app.zoomBy(1 / wheelZoomFactor(-100, 0));
    expect(app.zoom.value).toBeCloseTo(1, 2);
    expect(wheelZoomFactor(3, 1)).toBeCloseTo(wheelZoomFactor(48, 0)); // lines are 16px
    app.zoomBy(1000);
    expect(app.zoom.value).toBe(4);
    app.zoomBy(0.0001);
    expect(app.zoom.value).toBe(0.25);
  });

  it('opens a file at its saved zoom, and writes the zoom back without marking it modified', async () => {
    const app = await makeApp();
    expect(app.zoom.value).toBe(1);
    const zoomed = Workbook.createNew();
    zoomed.settings.zoom = 150;
    app.session.load(OdsWriter.write(zoomed, null), null, 'z.ods');
    expect(app.zoom.value).toBe(1.5);
    app.zoom.value = 0.8;
    expect(app.session.workbook.peek().settings.zoom).toBe(80);
    expect(app.session.modified.peek()).toBe(false);
    app.session.newDocument();
    expect(app.zoom.value).toBe(1);
    expect(app.session.workbook.peek().settings.zoom).toBeNull();
  });
});

describe('autofit columns', () => {
  it('fits rotated text by its turned box, even when it wraps', async () => {
    // A fake canvas: every character is 10px wide.
    Object.assign(document, {
      createElement: () => ({
        getContext: () => ({ font: '', measureText: (t: string) => ({ width: t.length * 10 }) }),
      }),
    });
    const app = await makeApp();
    const sheet = app.session.activeSheet.peek();
    for (const c of [0, 1, 2, 3]) sheet.setValue(0, c, CellValue.text('Quarterly total'));
    const turn = (column: number, rotation: number, wrapText = false) => {
      app.selection.select(0, column);
      app.style('Rotate', (s) => ({ ...s, alignment: { ...s.alignment, rotation, wrapText } }));
    };
    turn(1, 90);
    turn(2, 45);
    turn(3, 90, true);
    app.selection.select(0, 0);
    app.autofitColumns(CellRange.parse('A1:D1'));
    const [flat, up, slant, wrapped] = [0, 1, 2, 3].map((c) => sheet.getColumnWidth(c));
    expect(flat).toBe(150 + 2 * CELL_PAD_X);
    expect(up).toBeLessThan(40); // one line's height, not 150px of text
    expect(up).toBeGreaterThan(12);
    expect(slant).toBeGreaterThan(up!);
    expect(slant).toBeLessThan(flat!);
    expect(wrapped).toBe(up);
  });
});

describe('double-click on a column edge', () => {
  it('fits every selected whole column, ⌘-clicked ones included, or else just that column', async () => {
    Object.assign(document, {
      createElement: () => ({
        getContext: () => ({ font: '', measureText: (t: string) => ({ width: t.length * 10 }) }),
      }),
    });
    const app = await makeApp();
    const sheet = app.session.activeSheet.peek();
    for (const c of [0, 1, 2]) sheet.setValue(0, c, CellValue.text('a long heading'));
    const before = sheet.getColumnWidth(1);
    app.selection.selectRange({ top: 0, left: 0, bottom: MAX_ROW, right: 0 });
    app.selection.select(0, 2, { add: true });
    app.selection.select(MAX_ROW, 2, { extend: true });

    app.autofitColumn(2);
    expect(app.session.lastAction.peek()).toBe('Autofitted columns A:A, C:C');
    expect(sheet.getColumnWidth(0)).toBe(140 + 2 * CELL_PAD_X);
    expect(sheet.getColumnWidth(2)).toBe(140 + 2 * CELL_PAD_X);
    expect(sheet.getColumnWidth(1)).toBe(before);

    app.autofitColumn(1); // not selected: only its own column
    expect(sheet.getColumnWidth(1)).toBe(140 + 2 * CELL_PAD_X);
    expect(app.session.lastAction.peek()).toBe('Autofitted columns B:B');
  });
});

describe('paste at the sheet edge', () => {
  it('refuses text that would run past the last column instead of wrapping into the next row', async () => {
    const app = await makeApp();
    const last = SheetDimensions.maxColumns - 1;
    app.selection.select(3, last);
    app.paste('1\t2');

    const sheet = app.session.activeSheet.peek();
    expect(sheet.getValue(3, last)).toEqual(CellValue.empty);
    expect(sheet.getValue(4, 0)).toEqual(CellValue.empty);
    expect(app.toast.peek()?.text).toMatch(/doesn't fit/);
  });

  it('refuses its own copied block the same way', async () => {
    const app = await makeApp();
    const sheet = app.session.activeSheet.peek();
    app.session.execute(
      EditCommands.setInput(sheet.name, 0, 0, 'a', app.session.locale, app.session.workbook.peek()),
    );
    app.session.execute(
      EditCommands.setInput(sheet.name, 0, 1, 'b', app.session.locale, app.session.workbook.peek()),
    );
    app.selection.selectRange({ top: 0, left: 0, bottom: 0, right: 1 });
    app.copy(false, { setData: () => {} } as unknown as DataTransfer);
    const last = SheetDimensions.maxColumns - 1;
    app.selection.select(3, last);
    app.paste(null);

    const after = app.session.activeSheet.peek();
    expect(after.getValue(3, last)).toEqual(CellValue.empty);
    expect(after.getValue(4, 0)).toEqual(CellValue.empty);
    expect(app.toast.peek()?.text).toMatch(/doesn't fit/);
  });
});

describe('pasting text from other apps', () => {
  it('splits on tabs only: commas stay in the cell, lines become rows', async () => {
    const app = await makeApp();
    app.selection.select(0, 0);
    app.paste('Smith, John\nWe met, and talked.\n');
    const sheet = app.session.activeSheet.peek();
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('Smith, John'));
    expect(sheet.getValue(1, 0)).toEqual(CellValue.text('We met, and talked.'));
    expect(sheet.getValue(0, 1)).toEqual(CellValue.empty);
  });

  it('reads tab-separated rows, quoted cells with line breaks included', async () => {
    const app = await makeApp();
    app.selection.select(0, 0);
    app.paste('a,1\t"two\nlines"\r\nb\t3');
    const sheet = app.session.activeSheet.peek();
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('a,1'));
    expect(sheet.getValue(0, 1)).toEqual(CellValue.text('two\nlines'));
    expect(sheet.getValue(1, 0)).toEqual(CellValue.text('b'));
    expect(sheet.getValue(1, 1)).toEqual(CellValue.number(3));
  });

  it('reads one quoted cell, as a copy of a cell with a line break puts it', async () => {
    const app = await makeApp();
    app.selection.select(0, 0);
    app.paste('"say ""hi""\nthere"');
    expect(app.session.activeSheet.peek().getValue(0, 0)).toEqual(CellValue.text('say "hi"\nthere'));
  });
});

describe('clipboard across documents', () => {
  const dt = () => ({ setData: () => {} }) as unknown as DataTransfer;
  const odsWith = (text: string) => {
    const wb = Workbook.createNew();
    wb.worksheets[0]!.setValue(0, 0, CellValue.text(text));
    return OdsWriter.write(wb, null);
  };

  it('never moves a cut from one document into the next', async () => {
    const app = await makeApp();
    app.session.load(odsWith('from A'), ref('a.ods'));
    app.selection.select(0, 0);
    app.copy(true, dt());

    app.session.load(odsWith('kept in B'), ref('b.ods'));
    app.selection.select(2, 2);
    app.paste(null);

    const sheet = app.session.activeSheet.peek();
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('kept in B'));
    expect(sheet.getValue(2, 2)).toEqual(CellValue.empty);
  });

  it('still pastes a plain copy into another document', async () => {
    const app = await makeApp();
    app.session.load(odsWith('from A'), ref('a.ods'));
    app.selection.select(0, 0);
    app.copy(false, dt());

    app.session.load(odsWith('kept in B'), ref('b.ods'));
    app.selection.select(2, 2);
    app.paste(null);

    const sheet = app.session.activeSheet.peek();
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('kept in B'));
    expect(sheet.getValue(2, 2)).toEqual(CellValue.text('from A'));
  });
});

describe('slow opens', () => {
  /** A store whose reloads (Open Recent) resolve when the test says so. */
  function gated(app: Controller) {
    const fake = new FakeStore();
    withFake(app, fake);
    let release!: () => void;
    fake.reloadGate = new Promise<void>((r) => (release = r));
    const wb = Workbook.createNew();
    wb.worksheets[0]!.setValue(0, 0, CellValue.text('from the NAS'));
    fake.reloadBytes = OdsWriter.write(wb, null);
    (app as unknown as { nas: { open: (root: string, path: string) => Promise<OpenedDocument> } }).nas.open =
      () => fake.reload(ref('slow.ods'));
    return { release };
  }

  it('asks again before replacing edits made while the file downloaded', async () => {
    const app = await makeApp();
    const { release } = gated(app);
    const asked = vi.spyOn(app, 'confirmDiscard');
    const opening = app.openRecent({
      ref: { store: 'nas', id: 'root:/slow.ods', name: 'slow.ods' },
    } as never);
    await flush();
    app.session.activeSheet.peek().setValue(4, 4, CellValue.text('typed meanwhile'));
    app.session.touch();
    asked.mockResolvedValueOnce(false); // the user keeps their edits
    release();
    await opening;
    expect(asked).toHaveBeenCalledTimes(2);
    expect(app.session.activeSheet.peek().getValue(4, 4)).toEqual(CellValue.text('typed meanwhile'));
  });

  it('drops a slow open once New was chosen after it', async () => {
    const app = await makeApp();
    const { release } = gated(app);
    const opening = app.openRecent({
      ref: { store: 'nas', id: 'root:/slow.ods', name: 'slow.ods' },
    } as never);
    await flush();
    await app.newDocument();
    release();
    await opening;
    expect(app.session.document.peek()).toBeNull();
    expect(app.session.activeSheet.peek().getValue(0, 0)).toEqual(CellValue.empty);
  });
});

describe('deleting a sheet', () => {
  it('deletes an empty, unused sheet without asking', async () => {
    const app = await makeApp();
    app.session.addSheet();
    const { asked, stop } = answerConfirms(app);
    await app.deleteSheet(1);
    stop();
    expect(asked).toHaveLength(0);
    expect(app.session.sheetNames.peek()).toEqual(['Sheet1']);
  });

  it('asks for a sheet with data, without mentioning #REF! when nothing refers to it', async () => {
    const app = await makeApp();
    app.session.addSheet();
    app.session.workbook.peek().worksheets[1]!.setValue(0, 0, CellValue.number(1));
    const { asked, answer, stop } = answerConfirms(app);
    await app.deleteSheet(1);
    expect(asked).toHaveLength(1);
    expect(asked[0]).not.toMatch(/#REF!/);
    expect(app.session.sheetNames.peek()).toEqual(['Sheet1', 'Sheet2']);
    answer(true);
    await app.deleteSheet(1);
    stop();
    expect(app.session.sheetNames.peek()).toEqual(['Sheet1']);
  });

  it('mentions #REF! only when formulas elsewhere read the sheet', async () => {
    const app = await makeApp();
    app.session.addSheet();
    app.session.workbook.peek().worksheets[0]!.setFormula(0, 0, 'Sheet2.A1');
    const { asked, stop } = answerConfirms(app);
    await app.deleteSheet(1);
    stop();
    expect(asked[0]).toMatch(/1 formula on another sheet .*#REF!/);
  });
});

describe('closing the document', () => {
  function opened(app: Controller): void {
    const wb = Workbook.createNew();
    wb.worksheets[0]!.setValue(0, 0, CellValue.text('kept'));
    app.session.load(OdsWriter.write(wb, null), ref('a.ods'));
  }

  it('asks about unsaved changes and keeps the document when cancelled', async () => {
    const app = await makeApp();
    opened(app);
    app.session.touch();
    const closing = app.closeDocument();
    expect(app.dialog.peek()?.kind).toBe('confirm-discard');
    const d = app.dialog.peek() as Extract<DialogShown, { kind: 'confirm-discard' }>;
    d.resolve('cancel');
    await closing;
    expect(app.session.document.peek()?.name).toBe('a.ods');
    expect(app.session.modified.peek()).toBe(true);
  });

  it('goes back to a blank untitled workbook in the browser', async () => {
    const app = await makeApp();
    opened(app);
    app.session.touch();
    const closing = app.closeDocument();
    (app.dialog.peek() as Extract<DialogShown, { kind: 'confirm-discard' }>).resolve('discard');
    await closing;
    expect(app.session.document.peek()).toBeNull();
    expect(app.session.modified.peek()).toBe(false);
    expect(app.session.activeSheet.peek().getValue(0, 0)).toEqual(CellValue.empty);
  });

  it('closes a clean document without asking', async () => {
    const app = await makeApp();
    opened(app);
    await app.closeDocument();
    expect(app.dialog.peek()).toBeNull();
    expect(app.session.document.peek()).toBeNull();
  });

  it('closes the window on the desktop, after the same question', async () => {
    const app = await makeApp();
    const native = { confirmClose: vi.fn(), cancelClose: vi.fn() };
    (app as unknown as { native: typeof native }).native = native;
    opened(app);
    app.session.touch();
    const cancelled = app.closeDocument();
    (app.dialog.peek() as Extract<DialogShown, { kind: 'confirm-discard' }>).resolve('cancel');
    await cancelled;
    expect(native.cancelClose).toHaveBeenCalledOnce();
    expect(native.confirmClose).not.toHaveBeenCalled();

    const closing = app.closeDocument();
    (app.dialog.peek() as Extract<DialogShown, { kind: 'confirm-discard' }>).resolve('discard');
    await closing;
    expect(native.confirmClose).toHaveBeenCalledOnce();
    // The window goes; the document in it is not swapped for a blank one first.
    expect(app.session.document.peek()?.name).toBe('a.ods');
  });
});

describe('an open cell edit', () => {
  /** A grid with an edit open: committing it changes the document, unless the value is refused. */
  function editing(app: Controller, accept = true) {
    let open = true;
    const grid = {
      isEditing: () => open,
      commitEdit: vi.fn(() => {
        if (!accept) return false;
        if (open) app.session.touch();
        open = false;
        return true;
      }),
    };
    app.grid.current = grid as unknown as GridController;
    return grid;
  }

  it('commits before asking about unsaved changes, so the typed value is asked about', async () => {
    const app = await makeApp();
    app.session.load(OdsWriter.write(Workbook.createNew(), null), ref('a.ods'));
    const grid = editing(app);
    const asking = app.confirmDiscard();
    expect(grid.commitEdit).toHaveBeenCalledOnce();
    expect(app.dialog.peek()?.kind).toBe('confirm-discard');
    (app.dialog.peek() as Extract<DialogShown, { kind: 'confirm-discard' }>).resolve('cancel');
    expect(await asking).toBe(false);
  });

  it('keeps the document and the window when the typed value is refused', async () => {
    const app = await makeApp();
    const native = { confirmClose: vi.fn(), cancelClose: vi.fn() };
    (app as unknown as { native: typeof native }).native = native;
    app.session.load(OdsWriter.write(Workbook.createNew(), null), ref('a.ods'));
    editing(app, false);
    await app.requestClose();
    expect(app.dialog.peek()).toBeNull();
    expect(native.cancelClose).toHaveBeenCalledOnce();
    expect(native.confirmClose).not.toHaveBeenCalled();
  });

  it('counts as unsaved work when the tab closes', async () => {
    const app = await makeApp();
    expect(app.hasUnsavedWork).toBe(false);
    editing(app);
    expect(app.hasUnsavedWork).toBe(true);
  });
});

describe('printing the selection', () => {
  it('says so instead of printing when the selection holds no used cells', async () => {
    const app = await makeApp();
    app.session.activeSheet.peek().setValue(0, 0, CellValue.text('x'));
    app.selection.selectRange({ top: 5, left: 0, bottom: MAX_ROW, right: 0 });
    await app.print(true);
    expect(app.toast.peek()).toEqual({
      text: 'The selection is empty: there is nothing to print.',
      tone: 'info',
    });
  });
});

describe('LibreOffice locks', () => {
  it('warns on open when LibreOffice has the file, and opens it anyway', async () => {
    const app = await makeApp();
    app.adopt({
      ref: ref('shared.ods'),
      bytes: OdsWriter.write(Workbook.createNew(), null),
      lock: { name: 'Ada', date: '14.11.2026 10:22' },
    });
    expect(app.session.document.peek()?.name).toBe('shared.ods');
    expect(app.toast.peek()?.text).toBe(
      'Opened by Ada on 14.11.2026 10:22 in LibreOffice – your changes may conflict.',
    );
  });

  it('says nothing when there is no lock', async () => {
    const app = await makeApp();
    app.adopt({ ref: ref('free.ods'), bytes: OdsWriter.write(Workbook.createNew(), null) });
    expect(app.toast.peek()).toBeNull();
  });
});

describe('NAS file changed on the server', () => {
  function opened(app: Controller, fake: FakeStore): void {
    withFake(app, fake);
    app.session.load(OdsWriter.write(Workbook.createNew(), null), ref('a.ods'));
  }

  it('offers to reload a document without local edits, and reloads on yes', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    opened(app, fake);
    const wb = Workbook.createNew();
    wb.worksheets[0]!.setValue(0, 0, CellValue.text('their save'));
    fake.reloadBytes = OdsWriter.write(wb, null);
    expect(app.nasFileChanged()).toBe(true);
    const d = app.dialog.peek() as Extract<DialogShown, { kind: 'outside-change' }>;
    expect(d.kind).toBe('outside-change');
    expect(d.name).toBe('a.ods');
    d.resolve(true);
    await flush();
    await flush();
    expect(app.session.activeSheet.peek().getValue(0, 0)).toEqual(CellValue.text('their save'));
  });

  it('keeps the document when the offer is declined', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    opened(app, fake);
    const reload = vi.spyOn(app, 'reloadFromStore');
    app.nasFileChanged();
    (app.dialog.peek() as Extract<DialogShown, { kind: 'outside-change' }>).resolve(false);
    await flush();
    expect(reload).not.toHaveBeenCalled();
  });

  it('warns instead of asking when there are local edits; the save then meets the conflict', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    opened(app, fake);
    app.session.touch();
    expect(app.nasFileChanged()).toBe(true);
    expect(app.dialog.peek()).toBeNull();
    expect(app.toast.peek()?.text).toBe(
      '“a.ods” changed on the server. Saving will ask before overwriting it.',
    );
  });

  it('waits while a save is running or another dialog is open', async () => {
    const app = await makeApp();
    const fake = new FakeStore();
    opened(app, fake);
    app.dialog.value = { kind: 'about' };
    expect(app.nasFileChanged()).toBe(false);
    app.dialog.value = null;
    app.session.touch();
    const saving = app.save();
    await flush();
    expect(app.nasFileChanged()).toBe(false);
    fake.saves[0]!.finish();
    await saving;
  });
});
