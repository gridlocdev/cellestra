import { readFileSync } from 'node:fs';
import { CellValue, OdsWriter, Workbook } from '@cellestra/core';
import { signal } from '@preact/signals';
import { describe, expect, it, vi } from 'vitest';
import { WorkbookSession, type DocumentRef, type FileStore, type OpenedDocument } from '@cellestra/ui';
import { DocumentLifecycle } from '../src/app/documentLifecycle.ts';
import type { LifecycleHost } from '../src/app/lifecycleHost.ts';
import type { DocumentStores } from '../src/app/documentStores.ts';

/** A store whose saves finish (or fail) when the test says so. */
class SlowStore {
  readonly writes: { bytes: Uint8Array; finish: () => void; fail: (e: unknown) => void }[] = [];
  version = 1;
  reloadBytes = odsWith('on disk');

  save(ref: DocumentRef): Promise<DocumentRef> {
    return new Promise((resolve, reject) => {
      this.writes.push({
        bytes: new Uint8Array(),
        finish: () => resolve({ ...ref, version: `v${++this.version}` }),
        fail: reject,
      });
    });
  }
  async reload(ref: DocumentRef): Promise<OpenedDocument> {
    return { ref: { ...ref, version: 'v-disk' }, bytes: this.reloadBytes };
  }
}

function odsWith(text: string): Uint8Array {
  const wb = Workbook.createNew();
  wb.worksheets[0]!.setValue(0, 0, CellValue.text(text));
  return OdsWriter.write(wb, null);
}

const ref = (name: string): DocumentRef => ({ store: 'nas', name, id: `r:${name}`, version: 'v1' });
const flush = () => new Promise((r) => setTimeout(r, 0));

function setup(overrides: Partial<LifecycleHost> = {}) {
  const session = new WorkbookSession();
  const store = new SlowStore();
  const stores = {
    storeFor: () => store as unknown as FileStore,
    canSaveInPlace: () => true,
    primary: { saveAs: vi.fn(), pick: vi.fn() },
  } as unknown as DocumentStores;
  const native = { watch: vi.fn(), confirmClose: vi.fn(), cancelClose: vi.fn() };
  /** Set `desktop.on` to run as the Electron app. */
  const desktop = { on: false };
  const errors: string[] = [];
  const host: LifecycleHost = {
    session,
    stores,
    autoSaveSeconds: signal(0),
    native: () => (desktop.on ? native : null),
    questions: {
      discard: vi.fn(async () => 'discard' as const),
      csvLoss: vi.fn(async () => 'csv' as const),
      conflict: vi.fn(async () => 'cancel' as const),
    },
    commitOpenEdit: () => true,
    isEditing: () => false,
    dialogOpen: () => false,
    info: () => {},
    error: (e) => errors.push(e instanceof Error ? e.message : String(e)),
    recovery: { clear: vi.fn(async () => {}), keepAside: vi.fn(async () => {}) },
    remember: async () => {},
    replaced: () => {},
    reloadInPlace: (run) => run(),
    ...overrides,
  };
  const lifecycle = new DocumentLifecycle(host);
  const edit = (text: string) => {
    session.activeSheet.peek().setValue(3, 3, CellValue.text(text));
    session.touch();
  };
  return { session, store, host, lifecycle, native, desktop, errors, edit };
}

function openedDoc(name: string, text = name): OpenedDocument {
  return { ref: ref(name), bytes: odsWith(text) };
}

describe('saving', () => {
  it('runs one save at a time, and a second save of the same changes does not write again', async () => {
    const { lifecycle, store, edit, session } = setup();
    lifecycle.adopt(openedDoc('a.ods'));
    edit('changed');
    const first = lifecycle.save();
    const second = lifecycle.save();
    await flush();
    expect(store.writes).toHaveLength(1);
    expect(lifecycle.busy).toBe(true);
    store.writes[0]!.finish();
    expect(await first).toBe(true);
    expect(await second).toBe(true);
    expect(store.writes).toHaveLength(1);
    expect(session.modified.peek()).toBe(false);
    expect(lifecycle.busy).toBe(false);
  });

  it('stands auto-save down after a failed write, until a save succeeds', async () => {
    const { lifecycle, store, edit, errors, host } = setup();
    (host.autoSaveSeconds as ReturnType<typeof signal<number>>).value = 30;
    lifecycle.adopt(openedDoc('a.ods'));
    edit('one');
    const failing = lifecycle.autoSave();
    await flush();
    store.writes[0]!.fail(new Error('disk full'));
    expect(await failing).toBe(false);
    expect(errors).toEqual([
      "Auto-save couldn't write the file: disk full. Auto-save is paused until you save with ⌘S.",
    ]);

    edit('two');
    expect(await lifecycle.autoSave()).toBe(false);
    expect(store.writes).toHaveLength(1);

    const saving = lifecycle.save();
    await flush();
    store.writes[1]!.finish();
    await saving;
    edit('three');
    const resumed = lifecycle.autoSave();
    await flush();
    store.writes[2]!.finish();
    expect(await resumed).toBe(true);
  });
});

describe('opening', () => {
  it('lets a newer open win over a slower one', async () => {
    const { lifecycle, session } = setup();
    let releaseSlow!: (d: OpenedDocument) => void;
    const slow = lifecycle.open(() => new Promise((r) => (releaseSlow = r)));
    await flush();
    await lifecycle.open(async () => openedDoc('fast.ods'));
    releaseSlow(openedDoc('slow.ods'));
    await slow;
    expect(session.document.peek()?.name).toBe('fast.ods');
  });

  it('asks again about edits made while a slow open was on its way', async () => {
    const { lifecycle, session, host, edit } = setup();
    let release!: (d: OpenedDocument) => void;
    const opening = lifecycle.open(() => new Promise((r) => (release = r)));
    await flush();
    edit('typed meanwhile');
    vi.mocked(host.questions.discard).mockResolvedValueOnce('cancel');
    release(openedDoc('slow.ods'));
    await opening;
    expect(host.questions.discard).toHaveBeenCalledOnce();
    expect(session.activeSheet.peek().getValue(3, 3)).toEqual(CellValue.text('typed meanwhile'));
    expect(session.document.peek()).toBeNull();
  });

  it('runs `gone` and shows the error when the file no longer exists', async () => {
    const { lifecycle, errors } = setup();
    const gone = vi.fn(async () => {});
    await lifecycle.open(async () => {
      throw new DOMException('missing', 'NotFoundError');
    }, gone);
    expect(gone).toHaveBeenCalledOnce();
    expect(errors).toEqual(['missing']);
  });
});

describe('closing', () => {
  it('cancels the close when the user stays, and confirms it once changes are discarded', async () => {
    const { lifecycle, host, native, desktop, edit } = setup();
    desktop.on = true;
    lifecycle.adopt(openedDoc('a.ods'));
    edit('unsaved');
    vi.mocked(host.questions.discard).mockResolvedValueOnce('cancel');
    await lifecycle.requestClose();
    expect(native.cancelClose).toHaveBeenCalledOnce();
    expect(native.confirmClose).not.toHaveBeenCalled();

    await lifecycle.closeDocument();
    expect(native.confirmClose).toHaveBeenCalledOnce();
    expect(host.recovery.clear).toHaveBeenCalled();
  });
});

describe('Revert to Saved', () => {
  it('drops the recovery copy, since what is on screen is the file again', async () => {
    const { lifecycle, host, session, edit } = setup();
    lifecycle.adopt(openedDoc('a.ods'));
    edit('unsaved');
    vi.mocked(host.recovery.clear).mockClear();
    await lifecycle.reloadFromStore();
    expect(session.activeSheet.peek().getValue(0, 0)).toEqual(CellValue.text('on disk'));
    expect(host.recovery.clear).toHaveBeenCalled();
  });
});

describe('a file changed on disk', () => {
  it('waits while a dialog is open, then reloads a clean document', async () => {
    vi.useFakeTimers();
    try {
      let dialog = true;
      const { lifecycle, session } = setup({ dialogOpen: () => dialog });
      lifecycle.adopt(openedDoc('a.ods', 'before'));
      await lifecycle.externalChange();
      expect(session.activeSheet.peek().getValue(0, 0)).toEqual(CellValue.text('before'));
      dialog = false;
      await vi.advanceTimersByTimeAsync(1_000);
      expect(session.activeSheet.peek().getValue(0, 0)).toEqual(CellValue.text('on disk'));
    } finally {
      vi.useRealTimers();
    }
  });
});

describe('opening an .xlsx (PRD §8.19)', () => {
  const xlsx = readFileSync(
    new URL('../../../packages/core/test/serialization/fixtures/libreoffice-sample.xlsx', import.meta.url),
  );

  it('opens it as a new unsaved .ods, says so, and never watches or saves to the .xlsx', async () => {
    const notices: string[] = [];
    const { session, lifecycle, native, desktop } = setup({ info: (t) => notices.push(t) });
    desktop.on = true;
    lifecycle.adopt({
      ref: { store: 'native', name: 'Budget.xlsx', id: '/x/Budget.xlsx', version: 'v1' },
      bytes: xlsx,
    });
    await flush();
    expect(session.title.value).toBe('Budget.ods');
    expect(session.document.value).toBeNull();
    expect(session.importReport).not.toBeNull();
    expect(native.watch).toHaveBeenLastCalledWith(null);
    expect(notices[0]).toMatch(/^Opened Budget\.xlsx as a new workbook, Budget\.ods\./);
  });
});
