import { describe, expect, it } from 'vitest';
import { ChangePoller, type PollClock } from '../src/platform/changePoller.ts';
import { NasStore } from '../src/platform/nasStore.ts';

/** A clock the test advances by hand. */
function manualClock() {
  let now = 0;
  const timers = new Map<number, { at: number; fn: () => void }>();
  let nextId = 1;
  const clock: PollClock = {
    setTimeout: (fn, ms) => {
      const id = nextId++;
      timers.set(id, { at: now + ms, fn });
      return id;
    },
    clearTimeout: (id) => void timers.delete(id as number),
  };
  const advance = async (ms: number) => {
    now += ms;
    for (const [id, t] of [...timers]) {
      if (t.at > now) continue;
      timers.delete(id);
      t.fn();
    }
    await new Promise((r) => setTimeout(r, 0));
  };
  return { clock, advance, pending: () => timers.size };
}

function setup(initial = '"v1"') {
  const { clock, advance, pending } = manualClock();
  const state = { server: initial, known: initial as string | null, visible: true, accept: true };
  const reports: string[] = [];
  let fetches = 0;
  const poller = new ChangePoller({
    fetchVersion: () => {
      fetches++;
      return Promise.resolve(state.server);
    },
    knownVersion: () => state.known,
    onChange: (v) => {
      if (state.accept) reports.push(v);
      return state.accept;
    },
    intervalMs: 30_000,
    clock,
    isVisible: () => state.visible,
  });
  return { poller, state, reports, advance, pending, fetches: () => fetches };
}

describe('ChangePoller', () => {
  it('checks on the interval and reports a new version once', async () => {
    const t = setup();
    t.poller.start();
    await t.advance(29_999);
    expect(t.fetches()).toBe(0);
    await t.advance(1);
    expect(t.fetches()).toBe(1);
    expect(t.reports).toEqual([]);
    t.state.server = '"v2"';
    await t.advance(30_000);
    await t.advance(30_000);
    expect(t.reports).toEqual(['"v2"']);
    t.state.server = '"v3"';
    await t.advance(30_000);
    expect(t.reports).toEqual(['"v2"', '"v3"']);
  });

  it('reports again later when the app could not take the news', async () => {
    const t = setup();
    t.poller.start();
    t.state.server = '"v2"';
    t.state.accept = false;
    await t.advance(30_000);
    t.state.accept = true;
    await t.advance(30_000);
    expect(t.reports).toEqual(['"v2"']);
  });

  it('says nothing once the document caught up (saved or reloaded)', async () => {
    const t = setup();
    t.poller.start();
    t.state.server = '"v2"';
    t.state.known = '"v2"';
    await t.advance(30_000);
    expect(t.reports).toEqual([]);
  });

  it('drops a check that overlapped our own save', async () => {
    const t = setup();
    let release!: () => void;
    const gate = new Promise<void>((r) => (release = r));
    const poller = new ChangePoller({
      fetchVersion: async () => {
        await gate;
        return '"ours"';
      },
      knownVersion: () => t.state.known,
      onChange: (v) => (t.reports.push(v), true),
    });
    poller.start();
    const checking = poller.check();
    t.state.known = '"ours"'; // the save finished while the HEAD was out
    release();
    await checking;
    poller.stop();
    expect(t.reports).toEqual([]);
  });

  it('pauses while the page is hidden and catches up when it is shown', async () => {
    const t = setup();
    t.poller.start();
    t.state.visible = false;
    await t.advance(30_000);
    expect(t.fetches()).toBe(0);
    expect(t.pending()).toBe(0); // no timer runs while hidden
    t.state.server = '"v2"';
    t.state.visible = true;
    t.poller.wake();
    await t.advance(0);
    expect(t.reports).toEqual(['"v2"']);
    expect(t.pending()).toBe(1);
  });

  it('keeps polling through a failed request, and stops for good on stop()', async () => {
    const { clock, advance, pending } = manualClock();
    let calls = 0;
    const poller = new ChangePoller({
      fetchVersion: () => {
        calls++;
        return Promise.reject(new Error('offline'));
      },
      knownVersion: () => '"v1"',
      onChange: () => true,
      clock,
    });
    poller.start();
    await advance(30_000);
    await advance(30_000);
    expect(calls).toBe(2);
    poller.stop();
    expect(pending()).toBe(0);
    poller.wake();
    await advance(60_000);
    expect(calls).toBe(2);
  });

  it('treats a weakened ETag as the same version', async () => {
    const t = setup();
    t.poller.start();
    t.state.server = 'W/"v1"';
    await t.advance(30_000);
    expect(t.reports).toEqual([]);
  });
});

describe('NasStore.currentVersion', () => {
  it('asks with HEAD and returns the ETag, or null when the file is gone', async () => {
    const seen: { url: string; method: string | undefined }[] = [];
    const responses = [
      new Response(null, { headers: { ETag: '"e2"' } }),
      new Response(JSON.stringify({ error: 'Not found' }), { status: 404 }),
    ];
    const nas = new NasStore('/api', (url, init) => {
      seen.push({ url: String(url), method: init?.method });
      return Promise.resolve(responses.shift()!);
    });
    expect(await nas.currentVersion('docs:sub/a.ods')).toBe('"e2"');
    expect(await nas.currentVersion('docs:sub/a.ods')).toBeNull();
    expect(seen[0]).toEqual({ url: '/api/file?root=docs&path=sub%2Fa.ods', method: 'HEAD' });
  });
});
