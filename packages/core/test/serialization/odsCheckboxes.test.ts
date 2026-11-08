import { describe, expect, it } from 'vitest';
import { EditCommands } from '../../src/editing/editCommands.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { OdsPackage, readOds, writeOds } from '../../src/serialization/index.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { readEntry, roundTrip } from './helpers.ts';

function withCheckboxes(range = 'B2:B4', sheetName?: string) {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  if (sheetName) sheet.name = sheetName;
  sheet.setValue(2, 1, CellValue.boolean(true));
  const history = new EditHistory(workbook);
  history.execute(EditCommands.insertCheckboxes(workbook, sheet.name, CellRange.parse(range)));
  return { workbook, sheet, history };
}

const isCheckbox = (wb: Workbook, row: number, column: number) => {
  const sheet = wb.activeSheet;
  return !!wb.styles.get(sheet.effectiveStyleId(row, column)).checkbox;
};

describe('checkbox edit commands', () => {
  it('insert fills empty cells with FALSE, keeps values, and undoes', () => {
    const { workbook, sheet, history } = withCheckboxes();
    expect(sheet.getValue(1, 1)).toEqual(CellValue.boolean(false));
    expect(sheet.getValue(2, 1)).toEqual(CellValue.boolean(true));
    expect(isCheckbox(workbook, 3, 1)).toBe(true);
    history.undo();
    expect(sheet.find(1, 1)).toBeUndefined();
    expect(isCheckbox(workbook, 2, 1)).toBe(false);
  });

  it('setChecked writes booleans but leaves formulas', () => {
    const { workbook, sheet, history } = withCheckboxes();
    sheet.setFormula(3, 1, '=1=1');
    history.execute(
      EditCommands.setChecked(
        sheet.name,
        [
          [1, 1],
          [3, 1],
        ],
        true,
      ),
    );
    expect(sheet.getValue(1, 1)).toEqual(CellValue.boolean(true));
    expect(sheet.find(3, 1)!.formula).toBe('=1=1');
    history.execute(EditCommands.removeCheckboxes(workbook, sheet.name, CellRange.parse('B2:B4')));
    expect(isCheckbox(workbook, 1, 1)).toBe(false);
    expect(sheet.getValue(1, 1)).toEqual(CellValue.boolean(true));
  });
});

describe('checkboxes in .ods', () => {
  it('save as linked form:checkbox controls anchored in their cells', () => {
    const { workbook } = withCheckboxes('B2:B3', 'My Sheet');
    const content = readEntry(writeOds(workbook), 'content.xml');
    expect(content).toContain('<office:forms form:automatic-focus="false" form:apply-design-mode="false">');
    expect(content).toMatch(
      /<form:checkbox form:name="Checkbox B2" form:control-implementation="ooo:com.sun.star.form.component.CheckBox" xml:id="(cellestra-checkbox-1-1)" form:id="\1" form:linked-cell="'My Sheet'\.B2" form:current-state="unchecked"/,
    );
    expect(content).toContain('form:linked-cell="\'My Sheet\'.B3" form:current-state="checked"');
    // The control sits in the cell it is linked to, before the cell's text.
    expect(content).toMatch(
      /<table:table-cell [^>]*office:boolean-value="true"[^>]*>\s*<draw:control table:end-cell-address="'My Sheet'\.B3"[^>]*draw:control="cellestra-checkbox-1-2"\/>\s*<text:p>/,
    );
    // Forms come before the columns (ODF schema order).
    expect(content.indexOf('<office:forms')).toBeLessThan(content.indexOf('<table:table-column'));
  });

  it('round-trip back to the checkbox style, without piling up controls', () => {
    const { workbook } = withCheckboxes();
    const once = roundTrip(workbook);
    expect(isCheckbox(once, 1, 1)).toBe(true);
    expect(isCheckbox(once, 2, 1)).toBe(true);
    expect(isCheckbox(once, 0, 1)).toBe(false);
    expect(once.activeSheet.getValue(2, 1)).toEqual(CellValue.boolean(true));

    const content = readEntry(writeOds(roundTrip(once)), 'content.xml');
    expect(content.match(/<form:checkbox /g)).toHaveLength(3);
    expect(content.match(/<draw:control /g)).toHaveLength(3);
    expect(content.match(/<office:forms /g)).toHaveLength(1);
  });

  it('keeps a labelled checkbox and other controls as they were', () => {
    const { workbook } = withCheckboxes('B2');
    const bytes = writeOds(workbook);
    const content = readEntry(bytes, 'content.xml').replace(
      '</form:form>',
      '<form:button form:name="Go" form:id="b1" xml:id="b1" form:label="Go"/></form:form>',
    );
    const reread = readOds(replaceEntry(bytes, 'content.xml', content));
    expect(isCheckbox(reread, 1, 1)).toBe(true);
    const saved = readEntry(writeOds(reread), 'content.xml');
    expect(saved).toContain('form:name="Go"');
    expect(saved.match(/<form:checkbox /g)).toHaveLength(1);
    expect(saved.match(/<office:forms /g)).toHaveLength(1);
  });

  it('writes nothing for a sheet without checkboxes', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.setValue(0, 0, CellValue.boolean(true));
    expect(readEntry(writeOds(workbook), 'content.xml')).not.toContain('office:forms');
  });
});

function replaceEntry(bytes: Uint8Array, name: string, text: string): Uint8Array {
  const pkg = OdsPackage.read(bytes);
  pkg.setText(name, text);
  return pkg.write();
}
