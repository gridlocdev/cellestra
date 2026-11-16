import { after } from '../schedule.ts';

/** How often an open NAS document's version is checked while the page is visible. */
export const NAS_POLL_MS = 30_000;

/** The timer functions the poller uses; tests pass a hand-driven one. */
export interface PollClock {
  setTimeout(fn: () => void, ms: number): unknown;
  clearTimeout(handle: unknown): void;
}

export interface ChangePollerOptions {
  /** The file's version now (an ETag); null when it can't be told (gone, offline). May throw. */
  readonly fetchVersion: () => Promise<string | null>;
  /** The version the open document was read or last saved at; null stops the checks for now. */
  readonly knownVersion: () => string | null;
  /**
   * The file is now at `version`, not the known one. Returns false when the app can't say so yet
   * (a save or a question is under way): the next check reports it again.
   */
  readonly onChange: (version: string) => boolean;
  readonly intervalMs?: number;
  readonly clock?: PollClock;
  /** Whether the page is on screen. Hidden pages don't poll; `wake()` catches up when shown. */
  readonly isVisible?: () => boolean;
}

const systemClock: PollClock = {
  setTimeout: (fn, ms) => after(ms, fn),
  clearTimeout: (cancel) => (cancel as () => void)(),
};

/**
 * Watches a file that has no change notifications (a NAS file over HTTP) by asking for its version
 * on an interval. Each new version is reported once. A check that overlaps a save or reload of the
 * document (the known version moved while it ran) is dropped, so our own save is never reported.
 */
export class ChangePoller {
  readonly #options: ChangePollerOptions;
  readonly #clock: PollClock;
  readonly #intervalMs: number;
  #timer: unknown = null;
  #running = false;
  /** A tick found the page hidden and stopped; `wake()` restarts it. */
  #paused = false;
  #checking = false;
  /** The last version reported, so it is reported only once. */
  #reported: string | null = null;

  constructor(options: ChangePollerOptions) {
    this.#options = options;
    this.#clock = options.clock ?? systemClock;
    this.#intervalMs = options.intervalMs ?? NAS_POLL_MS;
  }

  start(): void {
    if (this.#running) return;
    this.#running = true;
    this.#schedule();
  }

  stop(): void {
    this.#running = false;
    this.#paused = false;
    if (this.#timer !== null) this.#clock.clearTimeout(this.#timer);
    this.#timer = null;
  }

  /** The page is visible again: check now if polling paused while it was hidden. */
  wake(): void {
    if (!this.#running || !this.#paused) return;
    this.#paused = false;
    void this.#tick();
  }

  /** One check: ask for the version and report it when it is new. */
  async check(): Promise<void> {
    if (this.#checking) return;
    const known = this.#options.knownVersion();
    if (known === null) return;
    this.#checking = true;
    let current: string | null;
    try {
      current = await this.#options.fetchVersion();
    } catch {
      return; // offline or refused: try again next time
    } finally {
      this.#checking = false;
    }
    if (!this.#running || current === null || this.#options.knownVersion() !== known) return;
    if (sameVersion(current, known)) {
      this.#reported = null;
      return;
    }
    if (this.#reported !== null && sameVersion(current, this.#reported)) return;
    if (this.#options.onChange(current)) this.#reported = current;
  }

  #schedule(): void {
    if (!this.#running) return;
    this.#timer = this.#clock.setTimeout(() => void this.#tick(), this.#intervalMs);
  }

  async #tick(): Promise<void> {
    this.#timer = null;
    if (!this.#running) return;
    if (this.#options.isVisible && !this.#options.isVisible()) {
      this.#paused = true;
      return;
    }
    await this.check();
    this.#schedule();
  }
}

/** ETags compare without a weak `W/` prefix a proxy may have added. */
function sameVersion(a: string, b: string): boolean {
  return a.replace(/^W\//, '') === b.replace(/^W\//, '');
}
