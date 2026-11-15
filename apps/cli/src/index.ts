import { launch } from './infra/workspace.ts';
import { renderHelp, type CommandSpec } from './spec.ts';
import { apply } from './commands/apply.ts';
import { cat, convert, evalCommand, get, info, recalc, set, sheets, validate } from './commands/baseline.ts';
import { chart, charts } from './commands/chart.ts';
import {
  cleanUp,
  dataValidation,
  deleteCols,
  deleteRows,
  hideCols,
  hideRows,
  hyperlink,
  insertCols,
  insertRows,
  merge,
  name,
  removeDuplicates,
  renameTable,
  showCols,
  showRows,
  sort,
  split,
  table,
  unmerge,
} from './commands/data.ts';
import {
  addSheet,
  completions,
  deleteSheet,
  diff,
  exportCommand,
  find,
  functions,
  importCommand,
  newCommand,
  open,
  put,
  renameSheet,
  stats,
} from './commands/extra.ts';
import { lint, trace } from './commands/inspect.ts';
import {
  clear,
  duplicateSheet,
  freeze,
  height,
  moveSheet,
  settings,
  style,
  width,
} from './commands/layout.ts';
import { copy, fill, filter, move, paint } from './commands/range.ts';
import { conditional, pageSetup, print } from './commands/printing.ts';
import { man, schema, skill } from './commands/reference.ts';

export const ROOT_DESCRIPTION =
  'Cellestra — a keyboard-friendly .ods spreadsheet toolkit for scripts and CI.';

/**
 * The command tree. `cellestra <file>` with no verb opens the workbook in the app; bare `cellestra`
 * prints help — this is a toolkit first and a launcher second.
 */
export function buildRoot(version: string): CommandSpec {
  const root: CommandSpec = {
    name: 'cellestra',
    description: ROOT_DESCRIPTION,
    args: [{ name: 'file', description: 'Workbook to open in Cellestra', arity: 'optional' }],
    subcommands: [
      // Baseline — the genuinely useful core.
      convert,
      cat,
      get,
      set,
      recalc,
      sheets,
      info,
      evalCommand,
      validate,
      trace,
      lint,
      // Extras.
      newCommand,
      importCommand,
      put,
      exportCommand,
      stats,
      find,
      diff,
      functions,
      addSheet,
      renameSheet,
      deleteSheet,
      duplicateSheet,
      moveSheet,
      apply,
      open,
      completions,
      man,
      schema,
      skill,
      // Data, structural, and layout edits — parity with the desktop app's Data/Insert menus.
      sort,
      filter,
      copy,
      move,
      fill,
      paint,
      cleanUp,
      removeDuplicates,
      split,
      table,
      renameTable,
      dataValidation,
      name,
      hyperlink,
      insertRows,
      deleteRows,
      insertCols,
      deleteCols,
      hideRows,
      showRows,
      hideCols,
      showCols,
      merge,
      unmerge,
      // Formatting and layout verbs — fronts on the matching apply ops.
      style,
      width,
      height,
      freeze,
      clear,
      settings,
      conditional,
      pageSetup,
      // Printing — the headless face of File ▸ Export PDF.
      print,
      // Chart sheets — the headless face of Insert ▸ Chart Sheet.
      chart,
      charts,
    ],
    run(p, con) {
      const path = p.argOrNull('file');
      if (path !== null) return launch(con, path);
      con.out(renderHelp(root, ['cellestra'], version));
      return 0;
    },
  };
  return root;
}
