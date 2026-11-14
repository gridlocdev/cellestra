// The reference texts for `apply`, the JSON reports and the conventions, kept in one place so `man`,
// `schema`, and the skill agree.

/** The `apply` document reference (`cellestra schema apply`). */
export const APPLY_REFERENCE = `apply — a batch of typed operations from JSON
=============================================

Document: {"ops": [ op, op, ... ]}   (or a bare array of ops)
Every op is an object with "op". Addresses: "Sheet1.B2" / "B2" (cell), "Sheet1.A1:C10" / "A1:C10" (range);
an address without a sheet means the op's "sheet" if given, else the active/first sheet.
Row/column spans: "3", "3:5"; "B", "B:D". Colors: "#RRGGBB". Sizes are points.

Cells
  {"op":"set", "cell":"B2", "value": 42}                   number | string | bool | null (null clears)
  {"op":"set", "cell":"B2", "value": "=SUM(B3:B9)"}        a string starting with = is a formula
  {"op":"set", "cell":"B2", "value": "2026-01-31"}         strings parse like typing: dates, numbers, %, bools
  {"op":"set", "at":"A1", "values": [["Item","Qty"],["Pen",3],["Ink","=B2*2"]]}   a 2-D block
  {"op":"clear", "range":"A1:C10", "what":"contents"}      contents (default) | formats | all

Formatting
  {"op":"style", "range":"A1:D1", "bold":true, "italic":false, "underline":false, "strikethrough":false,
                 "font":"Helvetica Neue", "size":12, "color":"#1A1A1A", "fill":"#EEEEEE",
                 "align":"center", "valign":"middle", "wrap":true,      align: general|left|center|right; valign: top|middle|bottom
                 "format":"#,##0.00", "border":"thin", "borderColor":"#000000",  border: none|thin|medium|thick|dashed|dotted|double
                 "rotation":90,     text angle in degrees (0-359, counter-clockwise; 0 removes it)
                 "checkbox":true}   a TRUE/FALSE checkbox (empty cells become FALSE; false removes it; saved
                                    as a linked form:checkbox, so LibreOffice shows it too)
                 Only the keys you give change; "fill": null removes a fill; "format":"General" resets.
  {"op":"conditional", "range":"B2:D20", "rules":[{"when":">", "value":100, "format":"light-red"},
                                                  {"when":"between", "value":1, "value2":10, "format":"yellow"},
                                                  {"when":"formula", "value":"=B2<AVERAGE($B$2:$B$20)", "format":"bold"}]}
                 rules replace the range's rules ([] removes them); "add":true appends to the existing ones.
                 when: > < >= <= = != between not-between formula; value/value2: a number, text or "=formula"
                 (relative to the range's top-left cell); format: light-red (default) | yellow | green | red-text | bold
  {"op":"width",  "columns":"B:D", "width": 120, "sheet":"Sheet1"}
  {"op":"height", "rows":"1", "height": 28}
  {"op":"freeze", "rows":1, "columns":0}                      0/0 unfreezes
  {"op":"merge", "range":"A1:D1"}      {"op":"unmerge", "cell":"A1"}
  {"op":"hide", "rows":"5:8"} / {"op":"hide", "columns":"C"}    {"op":"show", ...} likewise

Structure
  {"op":"add-sheet", "name":"Data", "at":1, "ifMissing":true}   at: 1-based position; ifMissing: no error if it exists
  {"op":"rename-sheet", "name":"Sheet1", "to":"Data"}   {"op":"delete-sheet", "name":"Old"}
  {"op":"duplicate-sheet", "name":"Sheet1", "as":"Copy"}   {"op":"move-sheet", "name":"Data", "to":1}
  {"op":"page-setup", "sheet":"Data", "paper":"a4|letter|legal|a3|a5|tabloid|210x297", "orientation":"portrait|landscape",
                      "margins":20 or {"top":20,"bottom":20,"left":15,"right":15},   millimetres
                      "scale":100  or  "fitWidth":1, "fitHeight":0  or  "pages":2,    0 = any number of pages that way
                      "gridlines":true, "headers":true, "center":"none|horizontal|vertical|both",
                      "order":"down-then-over|over-then-down", "reset":true}   only the keys you give change
  {"op":"settings", "locale":"en-US", "calc":"automatic|manual", "activeSheet":"Data"}
  {"op":"insert-rows", "at":3, "count":2}   {"op":"delete-rows", "at":3, "count":1}      (1-based)
  {"op":"insert-cols", "at":"B", "count":1} {"op":"delete-cols", "at":"B"}

Blocks
  {"op":"copy", "from":"A1:B3", "to":"Sheet2.D1", "mode":"all|values|formats|formulas"}   relative refs shift, like paste
  {"op":"move", "from":"A1:B3", "to":"D1"}                                              formulas travel unchanged; source cleared
  {"op":"fill", "range":"D2:D20", "from":"D2"}       repeat a source block down/across (default: the range's first row or column)
  {"op":"paint", "from":"A1:D1", "to":"A10:D20"}     format painter: copy formatting only, tiling the pattern
  {"op":"filter", "range":"A1:D100", "header":true, "columns": {"B": {"condition":">5"}, "C": {"values":["Red","Blue"]}}}
  {"op":"filter", "range":"A1:D100", "clear":true}   conditions are SUMIF-style (">5", "<>0", "*ink*"); rows are hidden, never deleted

Data
  {"op":"sort", "range":"A1:D100", "by":"B,A:desc", "header":true}
  {"op":"table", "range":"A1:D20", "name":"Sales", "header":true, "replace":true}   replace: redefine a table already on the range
  {"op":"table-options", "table":"Sales", "bandedRows":true, "headerColumn":true, "range":"A1:D30"}   table: a name or a cell in it
  {"op":"table-remove", "table":"Sales"}
  {"op":"validation", "range":"C2:C50", "list":["Red","Green"]}  or  "source":"Lists.A1:A8"
  {"op":"validation", "range":"B2:B50", "kind":"whole|decimal|date|text-length", "operator":"between|not-between|equal|not-equal|greater|less|greater-or-equal|less-or-equal",
                      "value1":1, "value2":100, "response":"stop|warning|information", "inputMessage":"…", "errorMessage":"…"}   a date bound may be "2026-01-31"
  {"op":"validation", "range":"B2:B50", "remove":true}
  {"op":"name", "name":"TaxRate", "range":"Config.B1"}   {"op":"unname", "name":"TaxRate"}
  {"op":"hyperlink", "cell":"A2", "url":"https://…", "tooltip":"…"}   "url": null removes it
  {"op":"chart", "type":"column", "title":"Sales", "from":"A1:C13",          or "values":["B2:B13"], "names":["Q1"], "categories":"A2:A13",
                 "stacking":"none", "legend":true, "xTitle":"Month", "yTitle":"USD"}   scatter: "xValues"; bubble: + "bubbleSizes"
  {"op":"chart-update", "sheet":"Chart1", "title":"…", "type":"line", …}   only the keys given change; data keys rebuild the series
  {"op":"chart-remove", "sheet":"Chart1"}

Run:  cellestra apply book.ods ops.json --in-place [--json] [--dry-run] [--locale en-US]
      cat ops.json | cellestra apply book.ods - -o out.ods --json
Ops apply in order; the first failing op aborts the whole batch and nothing is written.`;

/** The `--json` reply reference (`cellestra schema report`). */
export const REPORT_REFERENCE = `--json — the machine-readable reply
====================================

Cell (the one shape used everywhere a cell appears: get, cat.cells, find.matches, validate.errors, the edit report's changed/errors)
  {"cell": "Sheet1.B2",      always sheet-qualified
   "value": 12.5,            typed: number | string | true/false; absent when empty or an error
   "display": "12.50",       what the GUI shows (number format + locale)
   "formula": "=B3*C3",      only for formula cells, with the leading =
   "error": "#NAME?",        only when the value is an error
   "format": "0.00",         only when the number format is not General
   "style": {"bold": true, "fill": "#EEEEEE", "align": "center", "border": "thin"}}
                             only what differs from the default, using the same keys as apply's style op
                             (bold, italic, underline, strikethrough, font, size, color, fill, align, valign, wrap, rotation, border, borderColor, checkbox,
                              conditional: [{"when","value","value2","format"}…])

Edit report (set, apply, sort, chart, clean-up, … — anything that writes, with --json)
  {"ok": true,
   "file": "book.ods",        where it was written (-o or the input with --in-place)
   "dryRun": true,            only on --dry-run: nothing was written
   "changed": [Cell, …],      the cells the ops touched, with their values AFTER recalculation
                              (structural ops — insert/delete rows/cols/sheets, hide/show, freeze, width, height — list none)
   "recalculated": 12,        formula cells evaluated
   "errors": [Cell, …],       EVERY error cell now in the workbook — check this after each write
   "circular": ["Sheet1.A1"],
   "notes": ["…"],            the lines a human would have seen on stderr, when there are any
   …}                         verb-specific extras: applied, removed, sheet, table, sheets, range

Reads
  get       → Cell
  cat       → {"sheet", "range", "rows": [[display…]…], "cells": [Cell…]}   cells = populated only, row-major
  find      → {"term", "matches": [Cell…]}
  validate  → {"ok", "file", "opens", "errors": [Cell…], "circular": [addr…]}    exit 1 when !ok
  diff      → {"same", "differences", "addedSheets", "removedSheets", "cells": [{"cell","before","after"}…]}   exit 1 when !same
  eval      → {"ok", "formula", "value", "display", "error"}
  sheets    → [{"name","usedRange","rows","columns","cells","chart","frozenRows","frozenColumns",
                "tables":[{"name","range","header","headerColumn","bandedRows"}], "merges":["A1:D1"],
                "widths":{"A":120}, "heights":{"1":28}, "hiddenColumns":["C"], "hiddenRows":["5"],
                "validations":[{"range","kind","operator","value1","value2","list","source","response"}],
                "pageSetup":{"paper","orientation","margins","scale"|"fitWidth","fitHeight"|"pages","gridlines","headers","center","order"}}…]   (absent keys = default)
  info      → {"file", "sheets": […], "namedRanges": [{"name","target","scope"}…], "locale", "calculation", "populatedCells"}
  stats     → {"sheet","range","count","numeric","sum","average","min","max"}
  functions → [{"name","signature","category","returns","summary","arguments":[{"name","kind","required","repeatable","description"}],"example"}…]
  name list → [{"name","target","scope"}…]
  print     → {"output", "sheet", "pages", "format": "pdf"|"html"}
  man       → the whole command tree

Contract
  man --json carries "contract": 1 — the version of these shapes and of the apply document. Additive keys never bump it.

Errors
  any verb  → {"ok": false, "error": "…"} on stdout (exit 1), plus the usual one-line stderr message.`;

/** The conventions block at the top of `cellestra man`. */
export const CONVENTIONS = `## Conventions (read these first)

- **Addressing matches the app.** \`Sheet1.B2\` is a cell, \`Sheet1.A1:C10\` a range; omit the sheet
  (\`B2\`, \`A1:C10\`) to mean \`--sheet\` if given, else the active/first sheet. Rows are 1-based,
  columns are letters (\`B\`) or 1-based numbers. Spans: \`3:5\`, \`B:D\`.
- **Reads go to stdout**, in a table by default; \`--format csv|tsv|json\` for pipes. \`-\` is stdin/stdout.
- **Edits never overwrite without \`--in-place\` (\`-i\`) or \`-o <file>\`.** Every write is atomic (temp file
  + rename) and keeps unmodeled content — charts, pivots, anything Cellestra doesn't model — verbatim.
- **\`--json\` on any command** replaces the text with one JSON document. A write reports the cells it
  changed *and every error cell left in the workbook*, so a bad formula is visible in the same reply.
  \`cellestra schema report\` shows the shapes. A user error under \`--json\` is \`{"ok": false, "error"}\`.
- **\`apply\` is the batch verb**: one JSON document of typed ops (cells, blocks, styles, layout,
  structure, data, charts), one load, one recalculation, one save. \`cellestra schema apply\` lists the ops.
  Prefer it to a \`set\` per cell — it is faster and it fails as a whole, never half-written.
- **Check after you write**: \`validate\` (exit 1 on any error cell or circular reference), \`lint\` (the
  mistakes that aren't errors: numbers as text, short ranges, inconsistent formulas, rule violations),
  \`trace <cell>\` (what a formula reads, or \`--dependents\` who reads it), and \`diff\` (exit 1 on any
  difference in values, structure, or with \`--styles\` formatting) are made for scripts and CI.
  \`eval "=…"\` tests a formula with no file. A cell's \`style\` comes back in every \`--json\` reply.
- **Names are addresses**: a named range or a table name works wherever a range or cell does.
- **Same engine as the app**: evaluation, OpenFormula translation, and locale handling are identical;
  \`--locale en-US\` pins parsing/formatting for reproducible output.
- **No network, ever.** Only the files named on the command line are touched.
`;
