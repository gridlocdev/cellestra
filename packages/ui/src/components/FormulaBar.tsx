import { useEffect, useRef, useState } from 'preact/hooks';

const FIELD_LINE = 18;
const FIELD_PAD = 10;
const MAX_LINES = 8;

/** Name box (address / go to) + the active cell's content. Commits on Enter, reverts on Escape; ⌥↩ adds a line. */
export function FormulaBar(props: {
  readonly address: string;
  readonly content: string;
  readonly onCommit: (text: string) => void;
  readonly onGoTo: (reference: string) => boolean;
  readonly onDone: () => void;
  /** Nothing to show or edit (a chart sheet is active): both fields are blank and disabled. */
  readonly disabled?: boolean;
}) {
  const [text, setText] = useState(props.content);
  const [name, setName] = useState(props.address);
  const editing = useRef(false);
  const nameEditing = useRef(false);
  const [focused, setFocused] = useState(false);
  // A multi-line value shows its first line; while it is being edited the field opens downward
  // over the grid to show every line (up to a few, then it scrolls).
  const lines = focused && !props.disabled ? Math.min(MAX_LINES, text.split('\n').length) : 1;

  useEffect(() => {
    if (!editing.current) setText(props.content);
  }, [props.content]);
  useEffect(() => {
    if (!nameEditing.current) setName(props.address);
  }, [props.address]);

  return (
    <div class="cs-formula-bar">
      <input
        class="cs-name-box"
        aria-label="Cell reference"
        disabled={props.disabled}
        value={props.disabled ? '' : name}
        onFocus={(e) => {
          nameEditing.current = true;
          (e.target as HTMLInputElement).select();
        }}
        onBlur={() => {
          nameEditing.current = false;
          setName(props.address);
        }}
        onInput={(e) => setName((e.target as HTMLInputElement).value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            if (props.onGoTo(name.trim())) {
              nameEditing.current = false;
              props.onDone();
            }
          } else if (e.key === 'Escape') {
            nameEditing.current = false;
            setName(props.address);
            props.onDone();
          }
        }}
      />
      <span class="cs-fx" aria-hidden="true">
        fx
      </span>
      <div class="cs-formula-field">
        <textarea
          class={`cs-formula-input${lines > 1 ? ' cs-formula-input-open' : ''}`}
          aria-label="Cell contents"
          spellcheck={false}
          rows={1}
          disabled={props.disabled}
          value={props.disabled ? '' : text}
          style={{ height: `${FIELD_PAD + lines * FIELD_LINE}px` }}
          onFocus={() => {
            editing.current = true;
            setFocused(true);
          }}
          onBlur={() => {
            setFocused(false);
            if (editing.current && text !== props.content) props.onCommit(text);
            editing.current = false;
          }}
          onInput={(e) => setText((e.target as HTMLTextAreaElement).value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && (e.altKey || e.shiftKey)) {
              // ⌥↩ / ⇧↩ start a new line in the cell, as in the grid's own editor.
              e.preventDefault();
              const field = e.currentTarget;
              field.setRangeText('\n', field.selectionStart, field.selectionEnd, 'end');
              setText(field.value);
            } else if (e.key === 'Enter') {
              e.preventDefault();
              editing.current = false;
              if (text !== props.content) props.onCommit(text);
              props.onDone();
            } else if (e.key === 'Escape') {
              e.preventDefault();
              editing.current = false;
              setText(props.content);
              props.onDone();
            }
          }}
        />
      </div>
    </div>
  );
}
