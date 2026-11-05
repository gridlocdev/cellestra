import { CellRef } from './cellRef.ts';

const EMPTY: readonly CellRef[] = Object.freeze([]);

/**
 * Records which cells each formula reads (its precedents) and the reverse (its dependents), so the
 * recalculation engine recomputes only what an edit affects, in order. Edges are cell-granular.
 */
export class DependencyGraph {
  readonly #precedents = new Map<string, Map<string, CellRef>>();
  readonly #dependents = new Map<string, Map<string, CellRef>>();

  /**
   * Replace the set of cells `cell` reads. A self-edge (=A1+1 in A1, or =SUM(A1:A5) in A1) is kept so
   * the cell has in-degree 1 in the topological sort and is reported as circular.
   */
  setPrecedents(cell: CellRef, precedents: Iterable<CellRef>): void {
    this.#clearPrecedents(cell);
    const set = new Map<string, CellRef>();
    for (const p of precedents) set.set(CellRef.key(p), p);
    if (set.size === 0) return;
    const key = CellRef.key(cell);
    this.#precedents.set(key, set);
    for (const pk of set.keys()) {
      let dependents = this.#dependents.get(pk);
      if (!dependents) {
        dependents = new Map();
        this.#dependents.set(pk, dependents);
      }
      dependents.set(key, cell);
    }
  }

  /** Drop a cell's outgoing edges (when it stops being a formula). Incoming edges remain. */
  removeFormula(cell: CellRef): void {
    this.#clearPrecedents(cell);
  }

  getPrecedents(cell: CellRef): readonly CellRef[] {
    const set = this.#precedents.get(CellRef.key(cell));
    return set ? [...set.values()] : EMPTY;
  }

  getDependents(cell: CellRef): readonly CellRef[] {
    const set = this.#dependents.get(CellRef.key(cell));
    return set ? [...set.values()] : EMPTY;
  }

  /** Precedent keys of a cell (internal fast path). */
  precedentKeys(key: string): Iterable<string> {
    return this.#precedents.get(key)?.keys() ?? [];
  }

  /** Dependents of a cell by key (internal fast path). */
  dependentEntries(key: string): Iterable<[string, CellRef]> {
    return this.#dependents.get(key)?.entries() ?? [];
  }

  clear(): void {
    this.#precedents.clear();
    this.#dependents.clear();
  }

  #clearPrecedents(cell: CellRef): void {
    const key = CellRef.key(cell);
    const old = this.#precedents.get(key);
    if (!old) return;
    for (const pk of old.keys()) {
      const dependents = this.#dependents.get(pk);
      if (dependents) {
        dependents.delete(key);
        if (dependents.size === 0) this.#dependents.delete(pk);
      }
    }
    this.#precedents.delete(key);
  }
}
