import { readFileSync } from 'node:fs';
import { setHostLocale } from '@cellestra/core';
import { runCli } from './cli.ts';
import { launchApp } from './infra/launcher.ts';
import { processIo } from './infra/io.ts';

/** The `cellestra` executable: the real process streams, the system locale, and the app launcher. */

const pkg = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8')) as {
  version: string;
};
setHostLocale(Intl.DateTimeFormat().resolvedOptions().locale);
process.exitCode = runCli(process.argv.slice(2), processIo(launchApp), pkg.version);
