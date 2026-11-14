import { chmod, link, mkdtemp, mkdir, readFile, rm, stat, symlink, writeFile } from 'node:fs/promises';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import type { AddressInfo } from 'node:net';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { FileApi, parseRoots } from '../src/fileApi.ts';
import { bindAuthProblem, createServer, hostAllowed, parseHosts } from '../src/server.ts';

let dir: string;
let outside: string;
let base: string;
let close: () => void;

beforeAll(async () => {
  dir = await mkdtemp(path.join(os.tmpdir(), 'cellestra-root-'));
  outside = await mkdtemp(path.join(os.tmpdir(), 'cellestra-outside-'));
  await mkdir(path.join(dir, 'sub'));
  await writeFile(path.join(dir, 'a.ods'), 'AAA');
  await writeFile(path.join(dir, 'notes.txt'), 'hidden from the picker');
  await writeFile(path.join(dir, '.env'), 'SECRET=1');
  await mkdir(path.join(dir, '.git'));
  await writeFile(path.join(dir, '.git', 'x.ods'), 'hidden');
  await writeFile(path.join(outside, 'secret.ods'), 'nope');
  await symlink(outside, path.join(dir, 'escape'));
  const api = new FileApi(parseRoots(`Docs=${dir}`, '/'));
  const server = createServer({ api, staticDir: null, token: 't0k', maxBytes: 1024 });
  await new Promise<void>((r) => server.listen(0, '127.0.0.1', r));
  base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
  close = () => server.close();
});

afterAll(async () => {
  close();
  await rm(dir, { recursive: true, force: true });
  await rm(outside, { recursive: true, force: true });
});

const auth = { Authorization: 'Bearer t0k' };

describe('parseRoots', () => {
  it('slugs names into ids', () => {
    expect(parseRoots('My Docs=/a;Shared=/b', '/').map((r) => r.id)).toEqual(['my-docs', 'shared']);
  });
  it('rejects malformed entries', () => {
    expect(() => parseRoots('nopath', '/')).toThrow();
  });
});

describe('file API over HTTP', () => {
  it('requires the token', async () => {
    expect((await fetch(`${base}/api/roots`)).status).toBe(401);
    expect((await fetch(`${base}/api/roots`, { headers: { Authorization: 'Bearer wrong' } })).status).toBe(
      401,
    );
  });

  it('says without a token whether one is needed', async () => {
    const res = await fetch(`${base}/api/info`);
    expect(res.status).toBe(200);
    expect(await res.json()).toMatchObject({ auth: true });
  });

  it('lists only folders and spreadsheet files', async () => {
    const res = await fetch(`${base}/api/list?root=docs&path=`, { headers: auth });
    const names = ((await res.json()) as { name: string }[]).map((e) => e.name);
    expect(names).toContain('sub');
    expect(names).toContain('a.ods');
    expect(names).not.toContain('notes.txt');
  });

  it('refuses traversal and symlink escapes', async () => {
    expect(
      (await fetch(`${base}/api/file?root=docs&path=../../etc/passwd`, { headers: auth })).status,
    ).not.toBe(200);
    expect((await fetch(`${base}/api/file?root=docs&path=escape/secret.ods`, { headers: auth })).status).toBe(
      403,
    );
  });

  it('reads with an etag and saves with optimistic concurrency', async () => {
    const got = await fetch(`${base}/api/file?root=docs&path=a.ods`, { headers: auth });
    const etag = got.headers.get('etag')!;
    expect(await got.text()).toBe('AAA');

    const ok = await fetch(`${base}/api/file?root=docs&path=a.ods`, {
      method: 'PUT',
      headers: { ...auth, 'If-Match': etag },
      body: 'BBB',
    });
    expect(ok.status).toBe(200);
    expect(await readFile(path.join(dir, 'a.ods'), 'utf8')).toBe('BBB');

    const stale = await fetch(`${base}/api/file?root=docs&path=a.ods`, {
      method: 'PUT',
      headers: { ...auth, 'If-Match': etag },
      body: 'CCC',
    });
    expect(stale.status).toBe(412);
  });

  it('refuses to overwrite on create, and respects LibreOffice locks', async () => {
    const create = (name: string) =>
      fetch(`${base}/api/file?root=docs&path=${name}`, {
        method: 'PUT',
        headers: { ...auth, 'If-None-Match': '*' },
        body: 'x',
      });
    expect((await create('sub/new.ods')).status).toBe(200);
    expect((await create('sub/new.ods')).status).toBe(412);
    await writeFile(path.join(dir, '.~lock.locked.ods#'), '');
    expect((await create('locked.ods')).status).toBe(423);
  });

  it('rejects oversize uploads and non-spreadsheet names', async () => {
    const big = await fetch(`${base}/api/file?root=docs&path=big.ods`, {
      method: 'PUT',
      headers: auth,
      body: 'x'.repeat(2048),
    });
    expect(big.status).toBe(413);
    const txt = await fetch(`${base}/api/file?root=docs&path=evil.sh`, {
      method: 'PUT',
      headers: auth,
      body: 'x',
    });
    expect(txt.status).toBe(400);
  });
});

describe('read and write rules', () => {
  const put = (name: string, body: string, headers: Record<string, string> = {}) =>
    fetch(`${base}/api/file?root=docs&path=${name}`, {
      method: 'PUT',
      headers: { ...auth, ...headers },
      body,
    });

  it('reads only spreadsheet files outside dot folders', async () => {
    const get = (p: string) => fetch(`${base}/api/file?root=docs&path=${p}`, { headers: auth });
    expect((await get('.env')).status).toBe(403);
    expect((await get('notes.txt')).status).toBe(400);
    expect((await get('.git/x.ods')).status).toBe(403);
    expect((await put('.git/y.ods', 'x')).status).toBe(403);
  });

  it('lets only one of two concurrent saves from the same etag win', async () => {
    await writeFile(path.join(dir, 'race.ods'), 'v0');
    const etag = (await fetch(`${base}/api/file?root=docs&path=race.ods`, { headers: auth })).headers.get(
      'etag',
    )!;
    const results = await Promise.all(
      Array.from({ length: 5 }, (_, i) => put('race.ods', `v${i + 1}`, { 'If-Match': etag })),
    );
    const statuses = results.map((r) => r.status).sort();
    expect(statuses).toEqual([200, 412, 412, 412, 412]);
  });

  it('refuses to list dot folders', async () => {
    const list = (p: string) => fetch(`${base}/api/list?root=docs&path=${p}`, { headers: auth });
    expect((await list('.git')).status).toBe(403);
    expect((await list('sub/../.git')).status).toBe(403);
    expect((await list('sub')).status).toBe(200);
  });

  it('answers 400, not 500, for a folder named like a spreadsheet', async () => {
    await mkdir(path.join(dir, 'folder.ods'), { recursive: true });
    expect((await put('folder.ods', 'x')).status).toBe(400);
    expect((await fetch(`${base}/api/file?root=docs&path=folder.ods`, { headers: auth })).status).toBe(400);
    expect((await stat(path.join(dir, 'folder.ods'))).isDirectory()).toBe(true);
  });

  it('queues saves of one file under any spelling of its name', async () => {
    await writeFile(path.join(dir, 'Case.ods'), 'v0');
    const caseInsensitive = await stat(path.join(dir, 'CASE.ODS')).then(
      () => true,
      () => false,
    );
    const etag = (await fetch(`${base}/api/file?root=docs&path=Case.ods`, { headers: auth })).headers.get(
      'etag',
    )!;
    const names = caseInsensitive ? ['Case.ods', 'case.ods', 'CASE.ods', 'cAsE.ods'] : ['Case.ods'];
    const results = await Promise.all(
      names.flatMap((n, i) => [put(n, `a${i}`, { 'If-Match': etag }), put(n, `b${i}`, { 'If-Match': etag })]),
    );
    const statuses = results.map((r) => r.status).sort();
    expect(statuses.filter((s) => s === 200)).toHaveLength(1);
    expect(statuses.filter((s) => s === 412)).toHaveLength(statuses.length - 1);
  });

  it('answers If-Match on a missing file with 412', async () => {
    expect((await put('gone.ods', 'x', { 'If-Match': '"abc"' })).status).toBe(412);
  });

  it('keeps the file mode and hard links across a save', async () => {
    const file = path.join(dir, 'mode.ods');
    await writeFile(file, 'a');
    await chmod(file, 0o600);
    expect((await put('mode.ods', 'b')).status).toBe(200);
    expect((await stat(file)).mode & 0o777).toBe(0o600);

    const linked = path.join(dir, 'linked.ods');
    const other = path.join(outside, 'linked-copy.ods');
    await writeFile(linked, 'one');
    await link(linked, other);
    expect((await put('linked.ods', 'two')).status).toBe(200);
    expect(await readFile(other, 'utf8')).toBe('two');
  });
});

describe('static files', () => {
  let web: string;
  let port = 0;
  let stop = () => {};
  beforeAll(async () => {
    // The client lives under a folder that happens to be called "assets": only <root>/assets/ is immutable.
    web = path.join(await mkdtemp(path.join(os.tmpdir(), 'cellestra-web-')), 'assets', 'dist');
    await mkdir(path.join(web, 'assets'), { recursive: true });
    await writeFile(path.join(web, 'index.html'), '<!doctype html>');
    await writeFile(path.join(web, 'assets', 'app-123.js'), 'x');
    const api = new FileApi(parseRoots(`Docs=${dir}`, '/'));
    const server = createServer({ api, staticDir: web, token: null, maxBytes: 1024 });
    await new Promise<void>((r) => server.listen(0, '127.0.0.1', r));
    port = (server.address() as AddressInfo).port;
    stop = () => server.close();
  });
  afterAll(async () => {
    stop();
    await rm(path.dirname(path.dirname(web)), { recursive: true, force: true });
  });

  it('marks only the hashed bundles immutable', async () => {
    const index = await fetch(`http://127.0.0.1:${port}/`);
    expect(index.status).toBe(200);
    expect(index.headers.get('cache-control')).toBe('no-cache');
    const bundle = await fetch(`http://127.0.0.1:${port}/assets/app-123.js`);
    expect(bundle.headers.get('cache-control')).toContain('immutable');
  });

  it('answers a malformed percent-escape with 400', async () => {
    expect((await fetch(`http://127.0.0.1:${port}/%E0%A4%A`)).status).toBe(400);
  });
});

describe('Host check without a token', () => {
  let port = 0;
  let stop = () => {};
  beforeAll(async () => {
    const api = new FileApi(parseRoots(`Docs=${dir}`, '/'));
    const server = createServer({
      api,
      staticDir: null,
      token: null,
      maxBytes: 1024,
      allowedHosts: ['nas.lan'],
    });
    await new Promise<void>((r) => server.listen(0, '127.0.0.1', r));
    port = (server.address() as AddressInfo).port;
    stop = () => server.close();
  });
  afterAll(() => stop());

  // fetch() will not let us set Host, so use node:http directly.
  const statusFor = (host: string) =>
    new Promise<number>((resolve, reject) => {
      const req = http.get(
        { host: '127.0.0.1', port, path: '/api/roots', headers: { Host: host } },
        (res) => {
          res.resume();
          resolve(res.statusCode ?? 0);
        },
      );
      req.on('error', reject);
    });

  it('serves IP literals, localhost and CELLESTRA_HOSTS names, and refuses others with 421', async () => {
    expect(await statusFor(`127.0.0.1:${port}`)).toBe(200);
    expect(await statusFor(`[::1]:${port}`)).toBe(200);
    expect(await statusFor(`localhost:${port}`)).toBe(200);
    expect(await statusFor(`NAS.lan:${port}`)).toBe(200);
    expect(await statusFor(`evil.example.com:${port}`)).toBe(421);
  });

  it('parses the host list', () => {
    expect(parseHosts(' nas.lan , Other.Example:8787,')).toEqual(['nas.lan', 'other.example']);
    expect(hostAllowed(undefined, [])).toBe(false);
    expect(hostAllowed('10.0.0.5', [])).toBe(true);
    expect(hostAllowed('attacker.test', [])).toBe(false);
  });
});

describe('token requirement when binding', () => {
  it('lets a loopback address run without a token', () => {
    for (const host of ['127.0.0.1', '127.1.2.3', '::1', '[::1]', 'localhost', 'LOCALHOST'])
      expect(bindAuthProblem(host, null, false)).toBeNull();
  });

  it('refuses a non-loopback address without a token', () => {
    for (const host of ['0.0.0.0', '::', '192.168.1.10', 'nas.lan', ''])
      expect(bindAuthProblem(host, null, false)).toMatch(/CELLESTRA_TOKEN/);
  });

  it('allows a non-loopback address with a token, or with the reverse-proxy opt-out', () => {
    expect(bindAuthProblem('0.0.0.0', 'secret', false)).toBeNull();
    expect(bindAuthProblem('0.0.0.0', null, true)).toBeNull();
  });
});
