import path from 'node:path';
import { FileApi, parseRoots } from './fileApi.ts';
import { bindAuthProblem, createServer, defaultStaticDir, parseHosts } from './server.ts';

/**
 * Cellestra NAS server. Configuration is environment-only so it drops straight into Docker:
 *   CELLESTRA_ROOTS   Name=/dir;Other=/dir2   folders the client may open/save (required)
 *   CELLESTRA_STATIC  path to the built web client (default ../web/dist); empty = API only
 *   CELLESTRA_TOKEN   shared bearer token; required unless CELLESTRA_HOST is a loopback address
 *   CELLESTRA_NO_AUTH 1 = run on a non-loopback address without a token (a reverse proxy in front
 *                     must authenticate)
 *   CELLESTRA_HOST    bind address (default 0.0.0.0)
 *   CELLESTRA_PORT    port (default 8787); plain PORT still works when it is unset
 *   CELLESTRA_HOSTS   extra Host names (comma-separated) served when there is no token; IP
 *                     literals and localhost are always accepted, anything else gets 421
 *   CELLESTRA_MAX_MB  largest upload in MiB (default 64)
 */
const env = process.env;
const roots = parseRoots(env.CELLESTRA_ROOTS ?? '', process.cwd());
if (roots.length === 0) {
  console.error('Set CELLESTRA_ROOTS, e.g. CELLESTRA_ROOTS="Documents=/data/documents"');
  process.exit(1);
}
const staticDir =
  env.CELLESTRA_STATIC === ''
    ? null
    : env.CELLESTRA_STATIC !== undefined
      ? path.resolve(env.CELLESTRA_STATIC)
      : defaultStaticDir(import.meta.url);
const port = Number(env.CELLESTRA_PORT ?? env.PORT ?? 8787);
const host = env.CELLESTRA_HOST ?? '0.0.0.0';
const token = env.CELLESTRA_TOKEN || null;
const authProblem = bindAuthProblem(host, token, env.CELLESTRA_NO_AUTH === '1');
if (authProblem) {
  console.error(authProblem);
  process.exit(1);
}

const server = createServer({
  api: new FileApi(roots),
  staticDir,
  token,
  maxBytes: maxMb(env.CELLESTRA_MAX_MB) * 1024 * 1024,
  allowedHosts: parseHosts(env.CELLESTRA_HOSTS),
});

function maxMb(spec: string | undefined): number {
  const n = Number(spec);
  return spec && Number.isFinite(n) && n > 0 ? n : 64;
}

server.listen(port, host, () => {
  console.log(`Cellestra listening on http://${host}:${port}`);
  for (const r of roots) console.log(`  root '${r.name}' -> ${r.dir}`);
  if (!token) console.log('  no CELLESTRA_TOKEN set: only a reverse proxy or this machine should reach it');
});

for (const sig of ['SIGINT', 'SIGTERM'] as const) process.on(sig, () => server.close(() => process.exit(0)));
