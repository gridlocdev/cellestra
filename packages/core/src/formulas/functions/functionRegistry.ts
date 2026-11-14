import { dateTimeFunctions } from './dateTime.ts';
import type { FormulaFunction } from './formulaFunction.ts';
import type { FunctionMetadata } from './functionMetadata.ts';
import { logicalFunctions } from './logical.ts';
import { lookupFunctions } from './lookup.ts';
import { mathFunctions } from './math.ts';
import { subtotalFunctions } from './subtotal.ts';
import { textFunctions } from './text.ts';

/**
 * Names of functions that recompute every recalculation pass. Upper-case. SUBTOTAL is here because
 * hiding or filtering rows changes its result without changing any cell it reads.
 */
const VOLATILE: ReadonlySet<string> = new Set(['TODAY', 'NOW', 'SUBTOTAL']);

/**
 * The one catalog of functions: evaluation, autocomplete, the function picker, and the Function
 * Reference all read it, so they can never list different sets. Names match case-insensitively.
 */
export class FunctionRegistry {
  readonly #functions = new Map<string, FormulaFunction>();

  /** Functions that recompute every recalculation pass (TODAY, NOW, SUBTOTAL), upper-case. */
  static readonly volatileFunctions: ReadonlySet<string> = VOLATILE;

  static isVolatile(name: string): boolean {
    return VOLATILE.has(name.toUpperCase());
  }

  register(fn: FormulaFunction): void {
    this.#functions.set(fn.metadata.name.toUpperCase(), fn);
  }

  get(name: string): FormulaFunction | undefined {
    return this.#functions.get(name.toUpperCase());
  }

  has(name: string): boolean {
    return this.#functions.has(name.toUpperCase());
  }

  get count(): number {
    return this.#functions.size;
  }

  /** All function metadata, alphabetical (ordinal). */
  allMetadata(): FunctionMetadata[] {
    return [...this.#functions.values()]
      .map((f) => f.metadata)
      .sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));
  }

  /** The registry with the complete v1 library. */
  static createDefault(): FunctionRegistry {
    const registry = new FunctionRegistry();
    for (const fn of [
      ...mathFunctions(),
      ...subtotalFunctions(),
      ...logicalFunctions(),
      ...textFunctions(),
      ...lookupFunctions(),
      ...dateTimeFunctions(),
    ]) {
      registry.register(fn);
    }
    return registry;
  }
}
