import { mkdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { CliError } from '../infra/errors.ts';
import { CONTRACT_VERSION, writeReport } from '../infra/report.ts';
import { saveText } from '../infra/workspace.ts';
import { APPLY_REFERENCE, CONVENTIONS, REPORT_REFERENCE } from '../referenceTexts.ts';
import { isFlag, type ArgumentSpec, type CommandSpec, type OptionSpec } from '../spec.ts';
import { jsonOpt } from './options.ts';

/**
 * The whole CLI on one page. `man` renders every command, argument, option, and default from the
 * live command tree (so it cannot drift from `--help`), framed by the conventions that matter more
 * than any one verb. `schema` prints the reference for one of the JSON contracts.
 */

const EXIT_CODES: Record<string, string> = {
  '0': 'Success. For validate: no errors. For diff: identical. For find: at least one match.',
  '1': 'A problem was found: validate found errors; diff found differences; find found nothing; lint --strict found something; eval hit an error value. Also bad input (bad address, missing file, unknown sheet, refused overwrite, invalid JSON) — one `cellestra: …` line on stderr, and under --json a {"ok": false, "error"} document on stdout.',
  '2': 'A usage error: an unknown command or option, a missing argument, or an option value that does not parse. Same stderr line and --json document as 1, and the help goes to stdout.',
  'other / stack trace': 'A bug in Cellestra, not in your input.',
};

interface Entry {
  readonly path: string;
  readonly command: CommandSpec;
}

const ordinal = (a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0);

function* flatten(command: CommandSpec, path: string): Generator<Entry> {
  for (const child of [...(command.subcommands ?? [])].sort((a, b) => ordinal(a.name, b.name))) {
    const childPath = `${path} ${child.name}`;
    if ((child.subcommands ?? []).length > 0) yield* flatten(child, childPath);
    else yield { path: childPath, command: child };
  }
}

const isOptional = (a: ArgumentSpec): boolean => a.arity === 'optional';
const argName = (a: ArgumentSpec): string => (isOptional(a) ? `[${a.name}]` : `<${a.name}>`);

function defaultText(o: OptionSpec): string | null {
  if (o.default === undefined || isFlag(o) || o.default === '') return null;
  return o.default;
}

function optionSpelling(o: OptionSpec): string {
  const names = [...(o.aliases ?? [])].sort((a, b) => a.length - b.length);
  names.push(o.name);
  const spelled = names.join(', ');
  return isFlag(o) ? spelled : `${spelled} <value>`;
}

function usage(entry: Entry): string {
  let text = entry.path;
  for (const a of entry.command.args ?? []) text += ' ' + argName(a);
  if ((entry.command.options ?? []).length > 0) text += ' [options]';
  return text;
}

function renderMarkdown(root: CommandSpec): string {
  let sb = '# cellestra — reference\n\n';
  sb += root.description + '\n\n';
  sb +=
    'Generated from the live command tree by `cellestra man`; `cellestra <command> --help` shows any one command.\n\n';
  sb += CONVENTIONS + '\n';
  sb += '## Commands\n\n';
  const entries = [...flatten(root, 'cellestra')];
  for (const entry of entries) sb += `- \`${entry.path}\` — ${entry.command.description}\n`;
  sb += '\n';
  for (const entry of entries) {
    sb += `### ${entry.path}\n\n`;
    sb += entry.command.description + '\n\n';
    sb += '    ' + usage(entry) + '\n\n';
    for (const a of entry.command.args ?? []) sb += `- \`${argName(a)}\` — ${a.description}\n`;
    for (const o of entry.command.options ?? []) {
      sb += `- \`${optionSpelling(o)}\``;
      if (o.required) sb += ' (required)';
      sb += ` — ${o.description}`;
      const d = defaultText(o);
      if (d !== null) sb += ` [default: ${d}]`;
      sb += '\n';
    }
    sb += '\n';
  }
  sb += '## Exit codes\n\n';
  for (const [code, meaning] of Object.entries(EXIT_CODES)) sb += `- \`${code}\` — ${meaning}\n`;
  sb += '\n## JSON contracts\n\n```\n' + APPLY_REFERENCE + '\n```\n\n```\n' + REPORT_REFERENCE + '\n```\n';
  return sb;
}

export const man: CommandSpec = {
  name: 'man',
  description:
    'Print the complete reference — every command, option, convention, and JSON shape — as one page',
  options: [jsonOpt],
  run(p, con, root) {
    if (p.flag('--json')) {
      writeReport(con, {
        tool: 'cellestra',
        contract: CONTRACT_VERSION,
        description: root.description,
        commands: [...flatten(root, 'cellestra')].map((c) => ({
          name: c.path,
          description: c.command.description,
          arguments: (c.command.args ?? []).map((a) => ({
            name: a.name,
            description: a.description,
            optional: isOptional(a) ? true : null,
          })),
          options: (c.command.options ?? []).map((o) => ({
            name: o.name,
            aliases: (o.aliases ?? []).length > 0 ? o.aliases : null,
            description: o.description,
            required: o.required ? true : null,
            default: defaultText(o),
            flag: isFlag(o) ? true : null,
          })),
        })),
        exitCodes: EXIT_CODES,
        schemas: { apply: APPLY_REFERENCE, report: REPORT_REFERENCE },
      });
      return 0;
    }
    con.out(renderMarkdown(root));
    return 0;
  },
};

export const schema: CommandSpec = {
  name: 'schema',
  description: 'Print the reference for a JSON contract: the `apply` ops document, or the `--json` report',
  args: [{ name: 'contract', description: 'apply | report' }],
  run(p, con) {
    switch (p.arg('contract').trim().toLowerCase()) {
      case 'apply':
      case 'ops':
        con.outLine(APPLY_REFERENCE);
        return 0;
      case 'report':
      case 'json':
      case 'cell':
        con.outLine(REPORT_REFERENCE);
        return 0;
      default:
        throw new CliError('Known contracts: apply, report.');
    }
  },
};

/** The agent skill, shipped beside the source so an installed `cellestra` carries matching instructions. */
export function skillText(): string {
  return readFileSync(new URL('../../skill/SKILL.md', import.meta.url), 'utf8');
}

export const skill: CommandSpec = {
  name: 'skill',
  description: 'Print the agent skill for this tool, or --install it into ~/.claude/skills/cellestra/',
  options: [
    {
      name: '--install',
      description: 'Write the skill into the skills directory instead of printing it',
      kind: 'flag',
    },
    {
      name: '--dir',
      description:
        'Skills directory to install into (default: ~/.claude/skills; the skill goes in its own `cellestra/` folder)',
    },
  ],
  run(p, con) {
    const text = skillText();
    if (!p.flag('--install')) {
      con.out(text);
      return 0;
    }
    const root = p.str('--dir') ?? join(con.io.homeDir(), '.claude', 'skills');
    const folder = join(root, 'cellestra');
    mkdirSync(folder, { recursive: true });
    const path = join(folder, 'SKILL.md');
    saveText(con, text, path);
    con.errLine(`installed ${path}`);
    return 0;
  },
};
