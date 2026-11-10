/**
 * Child process for `cellestra print`: prints one HTML file to PDF with a headless Chromium-family
 * browser over the DevTools protocol, then exits. Run as
 *   node pdfPrinter.ts <browser> <page.html> <out.pdf> <profile-dir>
 * It is a separate process because the CLI's commands are synchronous and this is not. DevTools is
 * used instead of `--print-to-pdf`, which hangs in some browsers (Brave on macOS), and it honours the
 * page's CSS page size and margins exactly as the app's own PDF export does.
 */
import { spawn } from 'node:child_process';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

const [browser, page, out, profile] = process.argv.slice(2) as [string, string, string, string];
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
const fail = (message: string): never => {
  process.stderr.write(message + '\n');
  process.exit(1);
};

const child = spawn(
  browser,
  [
    '--headless',
    '--disable-gpu',
    '--no-first-run',
    '--no-default-browser-check',
    '--use-mock-keychain',
    '--password-store=basic',
    '--disable-extensions',
    '--disable-component-update',
    '--remote-debugging-port=0',
    `--user-data-dir=${profile}`,
    'about:blank',
  ],
  { stdio: 'ignore' },
);
child.on('error', (e) => fail(`Couldn't start the browser: ${e.message}`));
const timer = setTimeout(() => {
  child.kill();
  fail('The browser took too long to make the PDF.');
}, 90_000);

try {
  const portFile = join(profile, 'DevToolsActivePort');
  for (let i = 0; i < 200 && !existsSync(portFile); i++) await sleep(100);
  if (!existsSync(portFile)) fail("The browser didn't open its DevTools port.");
  const port = readFileSync(portFile, 'utf8').split('\n')[0]!.trim();
  const target = (await (
    await fetch(`http://127.0.0.1:${port}/json/new?about:blank`, { method: 'PUT' })
  ).json()) as { webSocketDebuggerUrl: string };

  const ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => {
    ws.onopen = resolve;
    ws.onerror = () => reject(new Error('DevTools connection failed'));
  });
  let id = 0;
  type Reply = {
    id?: number;
    method?: string;
    result?: Record<string, unknown>;
    error?: { message: string };
  };
  const pending = new Map<number, (m: Reply) => void>();
  let onLoad = () => {};
  const loaded = new Promise<void>((resolve) => (onLoad = resolve));
  ws.onmessage = (m) => {
    const d = JSON.parse(String(m.data)) as Reply;
    if (d.id !== undefined) pending.get(d.id)?.(d);
    else if (d.method === 'Page.loadEventFired') onLoad();
  };
  const send = (method: string, params: Record<string, unknown> = {}) =>
    new Promise<Record<string, unknown>>((resolve, reject) => {
      const i = ++id;
      pending.set(i, (m) => (m.error ? reject(new Error(m.error.message)) : resolve(m.result ?? {})));
      ws.send(JSON.stringify({ id: i, method, params }));
    });

  // Navigate only after Page events are on, so the load event can't be missed (the page is
  // self-contained: inline styles, no scripts, so loaded means laid out).
  await send('Page.enable');
  await send('Page.navigate', { url: pathToFileURL(page).href });
  await loaded;
  const pdf = await send('Page.printToPDF', { printBackground: true, preferCSSPageSize: true });
  writeFileSync(out, Buffer.from(pdf['data'] as string, 'base64'));
  ws.close();
} catch (e) {
  fail(`The browser couldn't make the PDF: ${(e as Error).message}`);
} finally {
  clearTimeout(timer);
  child.kill();
}
