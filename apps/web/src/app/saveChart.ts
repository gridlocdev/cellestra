/**
 * The rules for saving one open document, as a table. `DocumentLifecycle` keeps the state and runs the
 * effects; what happens next is decided here. `docs/architecture.md` draws this table, and a test checks
 * the drawing matches.
 *
 * A few things stay outside the table, as conditions on when the lifecycle sends an event:
 * - Saves take turns: `save` and `saveAs` arrive only when no write is running, so a second ⌘S waits
 *   for the first.
 * - `timer` arrives only when auto-save is on, hasn't failed since the last save, and the document can
 *   be written without a question (it has a file, the store can write it in place, a CSV loses nothing).
 * - An auto-save write never asks: a conflict ends it as `failed`, and auto-save waits for ⌘S.
 * - `closeLater` holds a close until no write is running, then sends `close` again.
 * - Opening another document or starting a new one starts a new table at `clean`.
 */

export type SaveState =
  /** The file has every edit, and nothing is being written or read. */
  | 'clean'
  /** Edits the file doesn't have; nothing is being written. */
  | 'dirty'
  /** Writing the file (⌘S, auto-save or Save As); nothing changed since the write started. */
  | 'saving'
  /** Writing the file; edits since the write started, so the document stays unsaved after it. */
  | 'savingDirty'
  /** Save As of a document with nothing unsaved; it stays saved however the write ends. */
  | 'copying'
  /** The file changed elsewhere; nothing is written until the user picks a conflict choice. */
  | 'conflict'
  /** Reading the file after a change elsewhere, to take it if nothing is edited meanwhile. */
  | 'refreshing'
  /** Reading the file to replace this document's unsaved edits with it (Revert to Saved, or Reload in a conflict). */
  | 'reloading'
  /** The save-changes question before closing the window. */
  | 'askingClose'
  /** The window closed. */
  | 'closed';

export type SaveEvent =
  /** The user changed the document. */
  | 'edit'
  /** ⌘S, or Save in the close question. */
  | 'save'
  /** Save As, or ⌘S when the document has no file the store can write in place. */
  | 'saveAs'
  /** Auto-save's countdown ran out. */
  | 'timer'
  /** The write finished. */
  | 'saved'
  /** The write failed, the file picker was cancelled, or the browser only downloaded the file. */
  | 'failed'
  /** The write found the file changed or locked elsewhere. */
  | 'versionChanged'
  /** Conflict choice: write this document over the file. */
  | 'overwrite'
  /** Conflict choice (Reload), or Revert to Saved: replace this document with the file. */
  | 'reload'
  /** Cancel in the conflict or close question. */
  | 'cancel'
  /** Discard in the close question. */
  | 'discard'
  /** The file changed on disk or on the server. */
  | 'refresh'
  /** A read of the file finished. */
  | 'readDone'
  /** A read of the file failed. */
  | 'readFailed'
  /** The user closes the window. */
  | 'close';

export type Effect =
  /** (Re)start auto-save's countdown. */
  | 'startTimer'
  /** Write the document to its file. */
  | 'writeFile'
  /** Ask where to save, then write the document there. */
  | 'pickFile'
  /** Ask the conflict question: overwrite, Save As, reload, or cancel. */
  | 'showConflict'
  /** Read the file. */
  | 'readFile'
  /** Ignore the reads in progress: the document changed after they started. */
  | 'dropRead'
  /** Take the file just read, unless it is the same version. */
  | 'takeFile'
  /** Replace the document with the file just read, dropping its edits. */
  | 'replaceWithFile'
  /** Say the file changed elsewhere and that saving will ask before overwriting it. */
  | 'warnChanged'
  /** Look at the change on disk again in a moment. */
  | 'retryLater'
  /** Ask whether to save the changes before closing. */
  | 'askClose'
  /** Close once no write is running. */
  | 'closeLater'
  /** Let the window close. */
  | 'closeWindow'
  /** Keep the window open. */
  | 'cancelClose';

export interface Transition {
  readonly from: SaveState;
  readonly on: SaveEvent;
  readonly to: SaveState;
  readonly effects: readonly Effect[];
}

/** Every allowed move. An event with no row for the state it arrives in changes nothing. */
export const SAVE_TRANSITIONS: readonly Transition[] = [
  { from: 'clean', on: 'edit', to: 'dirty', effects: ['startTimer'] },
  { from: 'clean', on: 'saveAs', to: 'copying', effects: ['pickFile'] },
  // Revert to Saved with nothing unsaved is a refresh: an edit while it reads wins.
  { from: 'clean', on: 'reload', to: 'refreshing', effects: ['readFile'] },
  { from: 'clean', on: 'refresh', to: 'refreshing', effects: ['readFile'] },
  { from: 'clean', on: 'close', to: 'closed', effects: ['closeWindow'] },

  { from: 'dirty', on: 'edit', to: 'dirty', effects: ['startTimer'] },
  { from: 'dirty', on: 'save', to: 'saving', effects: ['writeFile'] },
  { from: 'dirty', on: 'timer', to: 'saving', effects: ['writeFile'] },
  { from: 'dirty', on: 'saveAs', to: 'saving', effects: ['pickFile'] },
  { from: 'dirty', on: 'reload', to: 'reloading', effects: ['readFile'] },
  { from: 'dirty', on: 'refresh', to: 'dirty', effects: ['warnChanged'] },
  { from: 'dirty', on: 'close', to: 'askingClose', effects: ['askClose'] },

  { from: 'saving', on: 'edit', to: 'savingDirty', effects: [] },
  { from: 'saving', on: 'saved', to: 'clean', effects: [] },
  { from: 'saving', on: 'failed', to: 'dirty', effects: ['startTimer'] },
  { from: 'saving', on: 'versionChanged', to: 'conflict', effects: ['showConflict'] },
  { from: 'saving', on: 'refresh', to: 'saving', effects: ['retryLater'] },
  { from: 'saving', on: 'close', to: 'saving', effects: ['closeLater'] },

  { from: 'savingDirty', on: 'edit', to: 'savingDirty', effects: [] },
  { from: 'savingDirty', on: 'saved', to: 'dirty', effects: ['startTimer'] },
  { from: 'savingDirty', on: 'failed', to: 'dirty', effects: ['startTimer'] },
  { from: 'savingDirty', on: 'versionChanged', to: 'conflict', effects: ['showConflict'] },
  { from: 'savingDirty', on: 'refresh', to: 'savingDirty', effects: ['retryLater'] },
  { from: 'savingDirty', on: 'close', to: 'savingDirty', effects: ['closeLater'] },

  { from: 'copying', on: 'edit', to: 'savingDirty', effects: [] },
  { from: 'copying', on: 'saved', to: 'clean', effects: [] },
  { from: 'copying', on: 'failed', to: 'clean', effects: [] },
  { from: 'copying', on: 'refresh', to: 'copying', effects: ['retryLater'] },
  { from: 'copying', on: 'close', to: 'copying', effects: ['closeLater'] },

  { from: 'conflict', on: 'overwrite', to: 'saving', effects: ['writeFile'] },
  { from: 'conflict', on: 'saveAs', to: 'saving', effects: ['pickFile'] },
  { from: 'conflict', on: 'reload', to: 'reloading', effects: ['readFile'] },
  { from: 'conflict', on: 'cancel', to: 'dirty', effects: ['startTimer'] },
  { from: 'conflict', on: 'refresh', to: 'conflict', effects: ['retryLater'] },
  { from: 'conflict', on: 'close', to: 'conflict', effects: ['closeLater'] },

  // An edit during a refresh wins: the read is ignored, and the next save finds any conflict.
  { from: 'refreshing', on: 'edit', to: 'dirty', effects: ['dropRead', 'startTimer'] },
  { from: 'refreshing', on: 'saveAs', to: 'copying', effects: ['dropRead', 'pickFile'] },
  // Another change while reading: only the newest read counts.
  { from: 'refreshing', on: 'refresh', to: 'refreshing', effects: ['dropRead', 'readFile'] },
  { from: 'refreshing', on: 'readDone', to: 'clean', effects: ['takeFile'] },
  { from: 'refreshing', on: 'readFailed', to: 'clean', effects: [] },
  { from: 'refreshing', on: 'close', to: 'closed', effects: ['dropRead', 'closeWindow'] },

  // The user chose to drop this document's edits, so edits made while it reads are dropped too.
  { from: 'reloading', on: 'readDone', to: 'clean', effects: ['replaceWithFile'] },
  { from: 'reloading', on: 'readFailed', to: 'dirty', effects: ['startTimer'] },
  { from: 'reloading', on: 'close', to: 'reloading', effects: ['closeLater'] },

  { from: 'askingClose', on: 'save', to: 'saving', effects: ['writeFile', 'closeLater'] },
  { from: 'askingClose', on: 'saveAs', to: 'saving', effects: ['pickFile', 'closeLater'] },
  { from: 'askingClose', on: 'discard', to: 'closed', effects: ['closeWindow'] },
  { from: 'askingClose', on: 'cancel', to: 'dirty', effects: ['cancelClose', 'startTimer'] },
];

const NONE: readonly Effect[] = [];

/** Where `event` takes `state`, and what to do on the way. */
export function step(state: SaveState, event: SaveEvent): { to: SaveState; effects: readonly Effect[] } {
  const row = SAVE_TRANSITIONS.find((t) => t.from === state && t.on === event);
  return row ? { to: row.to, effects: row.effects } : { to: state, effects: NONE };
}

/** The table as a Mermaid state diagram, for `docs/architecture.md`. */
export function toMermaid(): string {
  const lines = ['stateDiagram-v2', '  [*] --> clean'];
  for (const t of SAVE_TRANSITIONS) {
    const effects = t.effects.length ? ` / ${t.effects.join(', ')}` : '';
    lines.push(`  ${t.from} --> ${t.to} : ${t.on}${effects}`);
  }
  return lines.join('\n');
}
