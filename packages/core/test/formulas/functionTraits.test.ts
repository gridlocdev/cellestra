import { describe, expect, it } from 'vitest';
import type { CellRef } from '../../src/dependencies/cellRef.ts';
import { ReferenceCollector } from '../../src/dependencies/referenceCollector.ts';
import { evaluateFormula } from '../../src/formulas/evaluation/evaluateFormula.ts';
import { defineFunction, withTraits } from '../../src/formulas/functions/formulaFunction.ts';
import { createMetadata } from '../../src/formulas/functions/functionMetadata.ts';
import { FunctionRegistry } from '../../src/formulas/functions/functionRegistry.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const parse = (text: string) => {
  const parsed = Parser.tryParse(text);
  if (!parsed.ok) throw new Error(`could not parse ${text}`);
  return parsed.expr;
};

function reads(text: string): string[] {
  const found = new Map<string, CellRef>();
  ReferenceCollector.collect(parse(text), 'Sheet1', () => null, found);
  return [...found.values()].map((r) => `${r.row},${r.column}`).sort();
}

describe('function traits', () => {
  it('marks TODAY, NOW and SUBTOTAL volatile, and nothing else in the library', () => {
    const volatile = FunctionRegistry.standard
      .allMetadata()
      .filter((m) => FunctionRegistry.standard.get(m.name)!.volatile)
      .map((m) => m.name);
    expect(volatile).toEqual(['NOW', 'SUBTOTAL', 'TODAY']);
    expect(ReferenceCollector.isVolatile(parse('1+ROUND(NOW(), 0)'))).toBe(true);
    expect(ReferenceCollector.isVolatile(parse('SUM(A1:A3)'))).toBe(false);
  });

  it('reads volatility from the registry it is given', () => {
    const registry = FunctionRegistry.createDefault();
    registry.register(
      withTraits(
        defineFunction(createMetadata('RANDISH', 'math', 'Test.', 'number', 'RANDISH()'), () =>
          CellValue.number(4),
        ),
        { volatile: true },
      ),
    );
    expect(ReferenceCollector.isVolatile(parse('RANDISH()'), registry)).toBe(true);
    expect(ReferenceCollector.isVolatile(parse('RANDISH()'))).toBe(false);
  });

  it("adds SUMIF's sum range, sized like the criteria range, to what the formula reads", () => {
    expect(reads('SUMIF(A1:A3, "x", B1)')).toEqual(['0,0', '0,1', '1,0', '1,1', '2,0', '2,1']);
  });

  it('lets IFERROR see an argument on a missing sheet as an error to replace', () => {
    const workbook = Workbook.createNew();
    const site = { workbook, sheet: workbook.activeSheet.name };
    expect(evaluateFormula(parse('IFERROR(SUM(Gone.A1:A2), 7)'), site)).toEqual(CellValue.number(7));
    const v = evaluateFormula(parse('SUM(Gone.A1:A2)'), site);
    expect(v.kind === 'error' && v.error.token).toBe('#REF!');
  });

  it('evaluates against the sheet it is given, with an injected clock', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.number(2));
    const clock = () => ({ year: 2024, month: 1, day: 2, hour: 0, minute: 0, second: 0, millisecond: 0 });
    expect(evaluateFormula(parse('A1*3'), { workbook, sheet: sheet.name })).toEqual(CellValue.number(6));
    expect(evaluateFormula(parse('DAY(TODAY())'), { workbook, sheet: sheet.name, clock })).toEqual(
      CellValue.number(2),
    );
  });
});
