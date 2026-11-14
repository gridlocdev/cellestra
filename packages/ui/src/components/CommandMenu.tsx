import { useEffect, useMemo, useRef, useState } from 'preact/hooks';
import { compact, type CommandCatalog } from '../commands/command.ts';
import { Icon } from './Icon.tsx';

/** ⌘K: search and run any command. Recent commands come first when the query is empty. */
export function CommandMenu({
  catalog,
  onClose,
}: {
  readonly catalog: CommandCatalog;
  readonly onClose: () => void;
}) {
  const [query, setQuery] = useState('');
  const [index, setIndex] = useState(0);
  const results = useMemo(() => catalog.search(query).slice(0, 60), [catalog, query]);
  const pick = (id: string | undefined) => {
    onClose();
    // Let the menu unmount (and focus return to the grid) before the command runs.
    if (id) queueMicrotask(() => catalog.run(id));
  };
  return (
    <ListPopup
      placeholder="Type a command…"
      query={query}
      onQuery={(q) => {
        setQuery(q);
        setIndex(0);
      }}
      index={index}
      onIndex={setIndex}
      count={results.length}
      onPick={(i) => pick(results[i]?.id)}
      onClose={onClose}
    >
      {results.map((c, i) => (
        <li
          key={c.id}
          aria-selected={i === index}
          aria-disabled={!catalog.isEnabled(c)}
          style={catalog.isEnabled(c) ? undefined : { opacity: 0.45 }}
          onPointerMove={() => setIndex(i)}
          onClick={() => pick(c.id)}
        >
          {c.icon ? <Icon name={c.icon} /> : <span class="cs-icon" />}
          <span>{c.title}</span>
          <span class="cs-group">{c.group}</span>
          {c.shortcutText && <span class="cs-kbd">{c.shortcutText}</span>}
        </li>
      ))}
    </ListPopup>
  );
}

/** ⌘⇧P-style quick pick (sheet switcher and friends). */
export function QuickPick<T>(props: {
  readonly items: readonly T[];
  readonly label: (item: T) => string;
  readonly placeholder: string;
  readonly onPick: (item: T) => void;
  readonly onClose: () => void;
}) {
  const [query, setQuery] = useState('');
  const [index, setIndex] = useState(0);
  // Same rule as the command palette: spaces don't count.
  const q = compact(query);
  const results = props.items.filter((i) => compact(props.label(i)).includes(q));
  const pick = (i: number) => {
    const item = results[i];
    props.onClose();
    if (item !== undefined) props.onPick(item);
  };
  return (
    <ListPopup
      placeholder={props.placeholder}
      query={query}
      onQuery={(v) => {
        setQuery(v);
        setIndex(0);
      }}
      index={index}
      onIndex={setIndex}
      count={results.length}
      onPick={pick}
      onClose={props.onClose}
    >
      {results.map((item, i) => (
        <li
          key={props.label(item)}
          aria-selected={i === index}
          onPointerMove={() => setIndex(i)}
          onClick={() => pick(i)}
        >
          {props.label(item)}
        </li>
      ))}
    </ListPopup>
  );
}

function ListPopup(props: {
  readonly placeholder: string;
  readonly query: string;
  readonly onQuery: (q: string) => void;
  readonly index: number;
  readonly onIndex: (i: number) => void;
  readonly count: number;
  readonly onPick: (i: number) => void;
  readonly onClose: () => void;
  readonly children: preact.ComponentChildren;
}) {
  // `autoFocus` only works at page load; focus the field once it is in the DOM.
  const input = useRef<HTMLInputElement>(null);
  useEffect(() => input.current?.focus(), []);
  const onKeyDown = (e: KeyboardEvent) => {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      props.onIndex(Math.min(props.count - 1, props.index + 1));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      props.onIndex(Math.max(0, props.index - 1));
    } else if (e.key === 'Enter') {
      e.preventDefault();
      props.onPick(props.index);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      props.onClose();
    } else if (e.key === 'Tab') {
      // The field is the popup's only stop: keep focus in it rather than let Tab leave the modal.
      e.preventDefault();
    }
  };
  return (
    <div
      class="cs-backdrop"
      onPointerDown={(e) => {
        if (e.target === e.currentTarget) props.onClose();
      }}
    >
      <div class="cs-dialog cs-palette" role="dialog" aria-modal="true">
        <input
          ref={input}
          value={props.query}
          placeholder={props.placeholder}
          onInput={(e) => props.onQuery((e.target as HTMLInputElement).value)}
          onKeyDown={onKeyDown}
          role="combobox"
          aria-expanded="true"
        />
        <ul role="listbox">{props.children}</ul>
      </div>
    </div>
  );
}
