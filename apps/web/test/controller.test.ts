import { CellRange, CellValue, EditCommands, OdsWriter, SheetDimensions, Workbook } from '@cellestra/core';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import { ConflictError, type DocumentRef, type FileStore, type OpenedDocument } from '@cellestra/ui';
import type { AppController as Controller } from '../src/app/controller.ts';

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
  reloadGate: Promise<void> = Promise.resolve();

  pick(): Promise<OpenedDocument | null> {
    return Promise.resolve(null);
  }
  async reload(ref: DocumentRef): Promise<OpenedDocument> {
    await this.reloadGate;
    return { ref, bytes: this.reloadBytes };
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
  saveAs(): Promise<DocumentRef | null> {
    return Promise.resolve(null);
  }
}

function withFake(app: Controller, fake: FakeStore): void {
  (app as unknown as { nas: FileStore }).nas = fake;
  app.error = (e: unknown) => {
    throw e;
  };
}

const flush = () => new Promise((r) => setTimeout(r, 0));
const ref = (name: string): DocumentRef => ({ store: 'nas', name, id: `r:${name}`, version: 'v1' });
const csvBytes = new TextEncoder().encode('a,b\n1,2\n');

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

describe('grid focus at launch', () => {
  it('focuses the grid when focus sits on the page, not when a field or dialog has it', async () => {
    const app = await makeApp();
    const body = {};
    const doc = globalThis.document as unknown as { body: unknown; activeElement: unknown };
    doc.body = body;
    const focus = vi.fn();
    app.grid.current = { focus, beginEdit: () => {}, isEditing: () => false, commitEdit: () => true };

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

describe('zoom', () => {
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
    expect(flat).toBe(150 + 12);
    expect(up).toBeLessThan(40); // one line's height, not 150px of text
    expect(up).toBeGreaterThan(12);
    expect(slant).toBeGreaterThan(up!);
    expect(slant).toBeLessThan(flat!);
    expect(wrapped).toBe(up);
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
  const withConfirm = (answer: boolean) => {
    const confirm = vi.fn((_: string) => answer);
    const g = globalThis as { window?: unknown };
    const before = g.window;
    g.window = { confirm };
    return { confirm, restore: () => (g.window = before) };
  };

  it('deletes an empty, unused sheet without asking', async () => {
    const app = await makeApp();
    app.session.addSheet();
    const { confirm, restore } = withConfirm(false);
    app.deleteSheet(1);
    restore();
    expect(confirm).not.toHaveBeenCalled();
    expect(app.session.sheetNames.peek()).toEqual(['Sheet1']);
  });

  it('asks for a sheet with data, without mentioning #REF! when nothing refers to it', async () => {
    const app = await makeApp();
    app.session.addSheet();
    app.session.workbook.peek().worksheets[1]!.setValue(0, 0, CellValue.number(1));
    const { confirm, restore } = withConfirm(false);
    app.deleteSheet(1);
    expect(confirm).toHaveBeenCalledOnce();
    expect(confirm.mock.calls[0]![0]).not.toMatch(/#REF!/);
    expect(app.session.sheetNames.peek()).toEqual(['Sheet1', 'Sheet2']);
    confirm.mockReturnValue(true);
    app.deleteSheet(1);
    restore();
    expect(app.session.sheetNames.peek()).toEqual(['Sheet1']);
  });

  it('mentions #REF! only when formulas elsewhere read the sheet', async () => {
    const app = await makeApp();
    app.session.addSheet();
    app.session.workbook.peek().worksheets[0]!.setFormula(0, 0, 'Sheet2.A1');
    const { confirm, restore } = withConfirm(false);
    app.deleteSheet(1);
    restore();
    expect(confirm.mock.calls[0]![0]).toMatch(/1 formula on another sheet .*#REF!/);
  });
});
