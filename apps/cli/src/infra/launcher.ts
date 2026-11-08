import { spawn } from 'node:child_process';
import { CliError } from './errors.ts';

/**
 * Open the Cellestra desktop app (with a document, or none). On macOS `open -a` reaches the
 * installed bundle, running or not; the document travels as an Apple Event, which the app handles.
 */
export function launchApp(path: string | null): number {
  if (process.platform !== 'darwin') {
    throw new CliError(
      'Opening a workbook in the app from the command line is only supported on macOS; open it from the desktop app.',
    );
  }
  const args = ['-a', 'Cellestra', ...(path === null ? [] : [path])];
  try {
    const child = spawn('open', args, { stdio: 'ignore', detached: true });
    child.on('error', () => undefined);
    child.unref();
  } catch (e) {
    throw new CliError(`Couldn't launch Cellestra: ${(e as Error).message}`);
  }
  return 0;
}
