import { readFileSync, writeSync } from 'node:fs';
import { homedir } from 'node:os';

/**
 * Everything the tool touches outside the workbook file: the three standard streams, the home
 * folder (for `skill --install`), and the app launcher. Production wires the real process; tests
 * capture the streams and stub the launcher, so the whole CLI runs in-process.
 */
export interface CliIo {
  /** Text to stdout. */
  stdout(text: string): void;
  /** Text to stderr. */
  stderr(text: string): void;
  /** Raw bytes to stdout (an .ods written to `-`). */
  stdoutBytes(bytes: Uint8Array): void;
  /** All of stdin, as bytes. */
  readStdin(): Uint8Array;
  /** The user's home folder. */
  homeDir(): string;
  /** Open a workbook in the Cellestra app; returns the exit code. */
  launch(path: string): number;
}

/** The output side of a run: writes through `CliIo`, line helpers included. */
export class Console {
  readonly io: CliIo;

  constructor(io: CliIo) {
    this.io = io;
  }

  out(text: string): void {
    this.io.stdout(text);
  }

  outLine(text = ''): void {
    this.io.stdout(text + '\n');
  }

  err(text: string): void {
    this.io.stderr(text);
  }

  errLine(text = ''): void {
    this.io.stderr(text + '\n');
  }

  stdinText(): string {
    return new TextDecoder().decode(this.io.readStdin());
  }
}

function writeAll(fd: number, data: string | Uint8Array): void {
  const bytes = typeof data === 'string' ? Buffer.from(data, 'utf8') : data;
  let offset = 0;
  while (offset < bytes.length) {
    try {
      offset += writeSync(fd, bytes, offset, bytes.length - offset);
    } catch (e) {
      // A non-blocking pipe can push back; retry until the reader drains it.
      if ((e as NodeJS.ErrnoException).code !== 'EAGAIN') throw e;
    }
  }
}

/** The real process streams. Writes are synchronous so `process.exit` never truncates output. */
export function processIo(launch: (path: string) => number): CliIo {
  return {
    stdout: (text) => writeAll(1, text),
    stderr: (text) => writeAll(2, text),
    stdoutBytes: (bytes) => writeAll(1, bytes),
    readStdin: () => readFileSync(0),
    homeDir: () => homedir(),
    launch,
  };
}
