import { buildRoot } from './index.ts';
import { EXIT_OK, EXIT_PROBLEM, EXIT_USAGE, isExpectedError, UsageError } from './infra/errors.ts';
import { Console, type CliIo } from './infra/io.ts';
import { toJson } from './infra/json.ts';
import { ParseError, parseArgs, renderHelp } from './spec.ts';

const TOOL = 'cellestra';

/** Parse and run one invocation; returns the exit code. A bug (not the user's) is rethrown. */
function runOnce(argv: readonly string[], io: CliIo, version: string): number {
  const root = buildRoot(version);
  const con = new Console(io);
  let invocation;
  try {
    invocation = parseArgs(root, argv);
  } catch (e) {
    if (e instanceof ParseError) {
      // A parse error: the error on stderr, the help on stdout.
      con.errLine(e.message);
      con.errLine();
      const path = [TOOL];
      let command = root;
      for (const token of argv) {
        const child = command.subcommands?.find((c) => c.name === token);
        if (!child) break;
        command = child;
        path.push(child.name);
      }
      con.out(renderHelp(command, path, command === root ? version : null));
      return EXIT_USAGE;
    }
    throw e;
  }
  try {
    switch (invocation.kind) {
      case 'help':
        con.out(
          renderHelp(invocation.command, invocation.path, invocation.command === root ? version : null),
        );
        return EXIT_OK;
      case 'version':
        con.outLine(version);
        return EXIT_OK;
      case 'run':
        return invocation.parsed.command.run!(invocation.parsed, con, root);
    }
  } catch (e) {
    if (isExpectedError(e)) {
      con.errLine(`${TOOL}: ${e.message}`);
      return e instanceof UsageError ? EXIT_USAGE : EXIT_PROBLEM;
    }
    throw e;
  }
}

/**
 * Whether this invocation asked for `--json`, read the way the parser reads it — so `--json=true`
 * and `--json true` count and `--json false` doesn't. When the command line doesn't parse, the last
 * `--json` token decides, so the parse error still comes back as JSON.
 */
function jsonRequested(argv: readonly string[], version: string): boolean {
  try {
    const invocation = parseArgs(buildRoot(version), argv);
    if (invocation.kind === 'run') return invocation.parsed.flag('--json');
  } catch (e) {
    if (!(e instanceof ParseError)) throw e;
  }
  let json = false;
  for (let i = 0; i < argv.length; i++) {
    const token = argv[i]!;
    if (token === '--') break;
    if (token === '--json') {
      const next = argv[i + 1];
      json = next === undefined || next.toLowerCase() !== 'false';
    } else if (token.startsWith('--json=')) json = token.slice('--json='.length).toLowerCase() !== 'false';
  }
  return json;
}

/**
 * The whole tool, in-process: tests call this with a capturing `CliIo`. Under `--json` a user error
 * is a JSON document on stdout too — the same `ok: false` shape every report uses — as well as the
 * one-line stderr message. Both streams are captured, since a parse error prints help to stdout.
 */
export function runCli(argv: readonly string[], io: CliIo, version = '0.0.0'): number {
  if (!jsonRequested(argv, version)) return runOnce(argv, io, version);

  const out: (string | Uint8Array)[] = [];
  let err = '';
  const captured: CliIo = {
    ...io,
    stdout: (text) => out.push(text),
    stdoutBytes: (bytes) => out.push(bytes),
    stderr: (text) => {
      err += text;
    },
  };
  const exit = runOnce(argv, captured, version);
  if (err.length > 0) io.stderr(err);
  if (exit !== 0 && err.length > 0) {
    const message = err.startsWith(`${TOOL}: `) ? err.slice(TOOL.length + 2) : err;
    const lines = [
      ...new Set(
        message
          .split('\n')
          .map((l) => l.trim())
          .filter((l) => l.length > 0),
      ),
    ];
    io.stdout(toJson({ ok: false, error: lines.join(' ') }) + '\n');
    return exit;
  }
  for (const chunk of out) {
    if (typeof chunk === 'string') io.stdout(chunk);
    else io.stdoutBytes(chunk);
  }
  return exit;
}
