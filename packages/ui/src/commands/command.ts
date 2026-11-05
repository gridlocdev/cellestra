import { formatShortcut, matches, parseShortcut, type Shortcut } from './shortcut.ts';

/** Menu groups — the native menu bar and the palette's section labels. */
export type CommandGroup = 'File' | 'Edit' | 'View' | 'Insert' | 'Format' | 'Data' | 'Sheet' | 'Help';

export interface CommandSpec {
  readonly id: string;
  readonly title: string;
  readonly description: string;
  readonly group: CommandGroup;
  readonly shortcut?: string;
  /** Unadvertised extra chords (⌘+ alongside ⌘=). */
  readonly alternates?: readonly string[];
  readonly icon?: string;
  /** Extra search words for the palette. */
  readonly keywords?: string;
  readonly run: () => void | Promise<void>;
  readonly enabled?: () => boolean;
  /** Toggle commands: whether it is on. */
  readonly active?: () => boolean;
  /** Hidden from the palette (e.g. the palette command itself). */
  readonly hidden?: boolean;
  /** Works while a text field (formula bar, inline editor) has focus. */
  readonly global?: boolean;
}

export interface Command extends CommandSpec {
  readonly shortcutKey: Shortcut | null;
  readonly alternateKeys: readonly Shortcut[];
  readonly shortcutText: string;
}

/**
 * The one command catalog: menus, toolbar, ⌘K palette, and key bindings all read it. Running a
 * command from anywhere goes through `run`, which also records it for the palette's recents.
 */
export class CommandCatalog {
  readonly all: readonly Command[];
  private readonly byId: Map<string, Command>;
  private readonly recent: string[] = [];

  constructor(specs: readonly CommandSpec[]) {
    this.all = specs.map((s) => {
      const key = s.shortcut ? parseShortcut(s.shortcut) : null;
      return {
        ...s,
        shortcutKey: key,
        alternateKeys: (s.alternates ?? []).map(parseShortcut),
        shortcutText: key ? formatShortcut(key) : '',
      };
    });
    this.byId = new Map(this.all.map((c) => [c.id, c]));
    if (this.byId.size !== this.all.length) throw new Error('Duplicate command id');
  }

  get(id: string): Command | undefined {
    return this.byId.get(id);
  }

  isEnabled(c: Command): boolean {
    return c.enabled?.() ?? true;
  }

  run(id: string): boolean {
    const c = this.byId.get(id);
    if (!c || !this.isEnabled(c)) return false;
    if (!c.hidden) {
      const at = this.recent.indexOf(id);
      if (at >= 0) this.recent.splice(at, 1);
      this.recent.unshift(id);
      this.recent.length = Math.min(this.recent.length, 5);
    }
    void Promise.resolve(c.run()).catch((e: unknown) => console.error(`command ${id} failed`, e));
    return true;
  }

  get recentIds(): readonly string[] {
    return this.recent;
  }

  /** The command a key event triggers, if any. `inText` = focus is in a text field. */
  forEvent(e: KeyboardEvent, inText: boolean): Command | undefined {
    return this.all.find(
      (c) =>
        (!inText || c.global) && [c.shortcutKey, ...c.alternateKeys].some((k) => k !== null && matches(k, e)),
    );
  }

  /** Palette search: every query word must appear in the title, group, or keywords; title-prefix first. */
  search(query: string): Command[] {
    const words = query.toLowerCase().split(/\s+/).filter(Boolean);
    const visible = this.all.filter((c) => !c.hidden);
    if (!words.length) {
      const recent = this.recent.map((id) => this.byId.get(id)!).filter((c) => c && !c.hidden);
      return [...recent, ...visible.filter((c) => !recent.includes(c))];
    }
    const scored: [Command, number][] = [];
    for (const c of visible) {
      const title = c.title.toLowerCase();
      const hay = `${title} ${c.group.toLowerCase()} ${c.keywords?.toLowerCase() ?? ''} ${c.description.toLowerCase()}`;
      if (!words.every((w) => hay.includes(w))) continue;
      const score = title.startsWith(words[0]!) ? 0 : title.includes(words[0]!) ? 1 : 2;
      scored.push([c, score]);
    }
    return scored.sort((a, b) => a[1] - b[1]).map(([c]) => c);
  }
}
