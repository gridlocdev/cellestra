import { useCallback, useEffect, useRef, useState } from 'preact/hooks';
import { Icon } from './Icon.tsx';
import { followFocus, ROVING_GROUP, rovingItem, rovingStep } from './roving.ts';

/** Scroll a tab into view within the strip, only as far as needed. */
const reveal = (tab: HTMLElement | null): void =>
  tab?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });

/** The active tab: in view, and holding the keyboard when an arrow key moved to it. */
const activeTab = (tab: HTMLElement | null): void => {
  reveal(tab);
  followFocus(tab);
};

export interface SheetTabsProps {
  readonly names: readonly string[];
  readonly active: number;
  /** `keyboard`: an arrow key moved along the strip, and the keyboard stays on the tabs. */
  readonly onSelect: (index: number, via?: 'keyboard') => void;
  readonly onAdd: () => void;
  /** Returns an error message to show, or null when the rename succeeded. */
  readonly onRename: (index: number, name: string) => string | null;
  readonly onMove: (from: number, to: number) => void;
  readonly onMenu: (index: number, x: number, y: number) => void;
}

/**
 * Sheet tabs: click to switch, double-click to rename, drag to reorder, right-click for more. One Tab
 * stop, the active sheet's tab; ← → Home End switch sheets from there.
 */
export function SheetTabs(props: SheetTabsProps) {
  const [renaming, setRenaming] = useState<number | null>(null);
  const [draft, setDraft] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [dragFrom, setDragFrom] = useState<number | null>(null);
  const input = useRef<HTMLInputElement>(null);
  useEffect(() => {
    if (renaming !== null) input.current?.select();
  }, [renaming]);
  // The strip narrows when the status text beside it grows; keep the active tab showing.
  const resizes = useRef<ResizeObserver | null>(null);
  const watchStrip = useCallback((el: HTMLDivElement | null) => {
    resizes.current?.disconnect();
    resizes.current = null;
    if (!el || typeof ResizeObserver === 'undefined') return;
    resizes.current = new ResizeObserver(() => reveal(el.querySelector('[aria-selected="true"]')));
    resizes.current.observe(el);
  }, []);

  const commit = () => {
    if (renaming === null) return;
    const err = draft.trim() === props.names[renaming] ? null : props.onRename(renaming, draft.trim());
    if (err) setError(err);
    else {
      setRenaming(null);
      setError(null);
    }
  };

  return (
    <div
      class="cs-tabs"
      role="tablist"
      aria-label="Sheets"
      ref={watchStrip}
      {...ROVING_GROUP}
      onKeyDown={(e) => {
        if (renaming !== null || !(e.target as HTMLElement).matches('.cs-tab')) return;
        const next = rovingStep(e, props.active, props.names.length, 'horizontal');
        if (next === undefined) return;
        e.preventDefault();
        if (next !== props.active) props.onSelect(next, 'keyboard');
      }}
    >
      {props.names.map((name, i) =>
        renaming === i ? (
          <input
            key={`edit-${i}`}
            ref={input}
            class="cs-tab-input"
            value={draft}
            title={error ?? undefined}
            style={error ? { borderColor: 'var(--cs-danger)' } : undefined}
            onInput={(e) => setDraft((e.target as HTMLInputElement).value)}
            onBlur={() => {
              setRenaming(null);
              setError(null);
            }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') commit();
              else if (e.key === 'Escape') {
                setRenaming(null);
                setError(null);
              }
              e.stopPropagation();
            }}
          />
        ) : (
          <button
            key={name}
            class="cs-tab"
            role="tab"
            aria-selected={i === props.active}
            // The active sheet's tab comes into view, also when there are more tabs than fit.
            {...rovingItem(i === props.active, activeTab)}
            draggable
            onClick={() => props.onSelect(i)}
            onDblClick={() => {
              setDraft(name);
              setRenaming(i);
            }}
            onContextMenu={(e) => {
              e.preventDefault();
              props.onMenu(i, e.clientX, e.clientY);
            }}
            onDragStart={() => setDragFrom(i)}
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => {
              e.preventDefault();
              if (dragFrom !== null && dragFrom !== i) props.onMove(dragFrom, i);
              setDragFrom(null);
            }}
          >
            {name}
          </button>
        ),
      )}
      <button class="cs-btn" title="Add sheet" aria-label="Add sheet" onClick={props.onAdd}>
        <Icon name="plus" />
      </button>
      {error && (
        <span class="cs-status" style={{ color: 'var(--cs-danger)' }}>
          {error}
        </span>
      )}
    </div>
  );
}
