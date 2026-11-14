import { describe, expect, it } from 'vitest';
import { ClipboardBlock } from '../../src/editing/clipboardBlock.ts';
import { FillCommand } from '../../src/editing/fillCommand.ts';
import { PasteBlockCommand } from '../../src/editing/pasteBlockCommand.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { OpenFormulaTranslator } from '../../src/formulas/openFormulaTranslator.ts';
import { readOds, writeOds } from '../../src/serialization/index.ts';
import { odsWith, readEntry } from './helpers.ts';

const ORIGINAL = 'of:=MYSTERY([.A1] @ [.B1]; {1;2})';
const MANY_OTHER_FORMULAS = 20_000;

describe('Unparsed formulas', () => {
  it('are saved exactly as read, however many other formulas were loaded since', () => {
    const bytes = odsWith(
      `<table:table table:name="S"><table:table-row><table:table-cell table:formula="${ORIGINAL}" office:value-type="float" office:value="1"/></table:table-row></table:table>`,
    );
    const workbook = readOds(bytes);
    // Other files read in the same session must not push this formula's original text out.
    for (let i = 0; i < MANY_OTHER_FORMULAS; i++)
      OpenFormulaTranslator.fromOpenFormula(`of:=OTHER${i}([.A1] @ 1)`);
    const content = readEntry(writeOds(workbook), 'content.xml');
    expect(content).toContain('table:formula="of:=MYSTERY([.A1] @ [.B1]; {1;2})"');
  });
});

describe('Copies of unparsed formulas', () => {
  it('keep the file text when pasted or filled', () => {
    const bytes = odsWith(
      `<table:table table:name="S"><table:table-row><table:table-cell table:formula="${ORIGINAL}" office:value-type="float" office:value="1"/></table:table-row></table:table>`,
    );
    const workbook = readOds(bytes);
    const sheet = workbook.activeSheet;
    new PasteBlockCommand(
      sheet.name,
      [0, 3],
      ClipboardBlock.capture(sheet, CellRange.parse('A1'), false),
    ).apply(workbook);
    new FillCommand(sheet.name, CellRange.parse('A1'), CellRange.parse('A1:A2')).apply(workbook);
    const content = readEntry(writeOds(workbook), 'content.xml');
    expect(content.split('table:formula="of:=MYSTERY([.A1] @ [.B1]; {1;2})"').length - 1).toBe(3);
  });
});
