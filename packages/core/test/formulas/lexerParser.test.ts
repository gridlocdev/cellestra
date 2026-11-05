import { describe, expect, it } from 'vitest';
import type { TokenType } from '../../src/formulas/lexing/lexer.ts';
import { Lexer } from '../../src/formulas/lexing/lexer.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';

describe('Lexer and Parser', () => {
  it('lexer tokenizes a function call', () => {
    const tokens = Lexer.tokenize('SUM(A1:A3, 5)');
    expect(tokens.map((t) => t.type)).toEqual([
      'name',
      'leftParen',
      'name',
      'colon',
      'name',
      'comma',
      'number',
      'rightParen',
      'endOfInput',
    ]);
  });

  it.each<[string, TokenType]>([
    ['<=', 'lessEqual'],
    ['>=', 'greaterEqual'],
    ['<>', 'notEqual'],
  ])('lexer reads compound operator %s', (source, expected) => {
    expect(Lexer.tokenize(source)[0]!.type).toBe(expected);
  });

  it('lexer decodes doubled quotes in strings', () => {
    const token = Lexer.tokenize('"a""b"')[0]!;
    expect(token.type).toBe('text');
    expect(token.text).toBe('a"b');
  });

  it('lexer keeps quoted sheet name with reference', () => {
    const token = Lexer.tokenize("'My Sheet'.A1")[0]!;
    expect(token.type).toBe('name');
    expect(token.text).toBe("'My Sheet'.A1");
  });

  it('parser strips leading equals', () => {
    expect(Parser.parse('=1+2').kind).toBe('binary');
  });

  it('parser respects precedence', () => {
    // 1 + 2 * 3 → 1 + (2 * 3)
    const expr = Parser.parse('1 + 2 * 3');
    if (expr.kind !== 'binary') throw new Error('expected binary');
    expect(expr.operator).toBe('add');
    const right = expr.right;
    if (right.kind !== 'binary') throw new Error('expected binary');
    expect(right.operator).toBe('multiply');
  });

  it('parser classifies name kinds', () => {
    expect(Parser.parse('A1').kind).toBe('reference');
    expect(Parser.parse('A1:B2').kind).toBe('range');
    expect(Parser.parse('TaxRate').kind).toBe('name');
    expect(Parser.parse('SUM(A1)').kind).toBe('call');
    expect(Parser.parse('TRUE').kind).toBe('literal');
  });

  it('parser reports position on error', () => {
    const result = Parser.tryParse('1 + ');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.position).toBe(4);
  });

  it('parser rejects unbalanced parens', () => {
    expect(Parser.tryParse('SUM(A1').ok).toBe(false);
  });
});
