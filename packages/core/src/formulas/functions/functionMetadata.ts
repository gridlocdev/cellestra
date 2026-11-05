/** The category a function belongs to (drives the Formulas menus and the Function Reference). */
export type FunctionCategory = 'math' | 'statistical' | 'logical' | 'text' | 'lookup' | 'dateTime';

/** The data shape a function argument expects (for autocomplete hints / docs). */
export type ArgumentType = 'any' | 'number' | 'text' | 'boolean' | 'range';

/** One argument in a function's signature. */
export interface FunctionArgument {
  readonly name: string;
  readonly type: ArgumentType;
  readonly optional: boolean;
  readonly repeatable: boolean;
  readonly description?: string;
}

/**
 * Everything three consumers need about a function so they never disagree: the evaluator (arity),
 * autocomplete (hints), and the Function Reference (docs).
 */
export interface FunctionMetadata {
  readonly name: string;
  readonly category: FunctionCategory;
  readonly summary: string;
  readonly args: readonly FunctionArgument[];
  readonly returnType: ArgumentType;
  readonly example: string;
  /** Number of required arguments. */
  readonly minArgs: number;
  /** Maximum arguments; `Infinity` when the last argument repeats. */
  readonly maxArgs: number;
  /** A one-line signature like `SUM(number, …)` / `ROUND(number, [digits])`. */
  readonly signature: string;
}

export function createMetadata(
  name: string,
  category: FunctionCategory,
  summary: string,
  returnType: ArgumentType,
  example: string,
  ...args: FunctionArgument[]
): FunctionMetadata {
  const minArgs = args.filter((a) => !a.optional).length;
  const maxArgs = args.length > 0 && args[args.length - 1]!.repeatable ? Infinity : args.length;
  const parts = args.map((a) => {
    const text = a.repeatable ? `${a.name}, …` : a.name;
    return a.optional ? `[${text}]` : text;
  });
  return Object.freeze({
    name,
    category,
    summary,
    args: Object.freeze(args),
    returnType,
    example,
    minArgs,
    maxArgs,
    signature: `${name}(${parts.join(', ')})`,
  });
}

export function acceptsArgCount(metadata: FunctionMetadata, count: number): boolean {
  return count >= metadata.minArgs && count <= metadata.maxArgs;
}

interface ArgOptions {
  readonly optional?: boolean;
  readonly repeatable?: boolean;
}

const arg =
  (type: ArgumentType) =>
  (name: string, options: ArgOptions = {}): FunctionArgument =>
    Object.freeze({
      name,
      type,
      optional: options.optional ?? false,
      repeatable: options.repeatable ?? false,
    });

/** Argument constructors by type. */
export const Arg = {
  any: arg('any'),
  number: arg('number'),
  text: arg('text'),
  boolean: arg('boolean'),
  range: arg('range'),
} as const;
