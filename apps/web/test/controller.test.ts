import { CellValue, EditCommands, OdsWriter, Workbook } from '@cellestra/core';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import type { DocumentRef, FileStore, OpenedDocument } from '@cellestra/ui';
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
    app.grid.current = { focus, beginEdit: () => {}, isEditing: () => false };

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
