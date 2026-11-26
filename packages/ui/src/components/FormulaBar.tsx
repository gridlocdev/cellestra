import { useEffect, useRef, useState } from 'preact/hooks';
import { FormulaBarDraft } from './formulaBarDraft.ts';
import { insertLineBreak } from './lineBreak.ts';

const FIELD_LINE = 18;
const FIELD_PAD = 10;
const MAX_LINES = 8;

/** Name box (address / go to) + the active cell's content. Commits on Enter, reverts on Escape; ⌥↩ adds a line. */
export function FormulaBar(props: {
  readonly address: string;
  readonly content: string;
  /** Identifies the active cell in its sheet and document; typing never carries over to another. */
  readonly target: string;
  readonly onCommit: (text: string) => void;
  readonly onGoTo: (reference: string) => boolean;
  readonly onDone: () => void;
  /** Nothing to show or edit (a chart sheet is active): both fields are blank and disabled. */
  readonly disabled?: boolean;
  /** Filled with the Name Box, so Go To can focus it. */
  readonly nameBoxRef?: { current: HTMLInputElement | null };
}) {
  const draft = useRef<FormulaBarDraft | null>(null);
  draft.current ??= new FormulaBarDraft(props.content);
  const [text, setText] = useState(props.content);
  const [name, setName] = useState(props.address);
  const nameEditing = useRef(false);
  const [focused, setFocused] = useState(false);
  // A multi-line value shows its first line; while it is being edited the field opens downward
  // over the grid to show every line (up to a few, then it scrolls).
  const lines = focused && !props.disabled ? Math.min(MAX_LINES, text.split('\n').length) : 1;

  useEffect(() => {
    draft.current!.sync(props.content, props.target);
    setText(draft.current!.text);
  }, [props.content, props.target]);
  useEffect(() => {
    if (!nameEditing.current) setName(props.address);
  }, [props.address]);

  return (
    <div class="cs-formula-bar">
      <input
        ref={(el) => {
          if (props.nameBoxRef) props.nameBoxRef.current = el;
        }}
        class="cs-name-box cs-focus-inset"
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
          class={`cs-formula-input cs-focus-inset${lines > 1 ? ' cs-formula-input-open' : ''}`}
          aria-label="Cell contents"
          spellcheck={false}
          rows={1}
          disabled={props.disabled}
          value={props.disabled ? '' : text}
          style={{ height: `${FIELD_PAD + lines * FIELD_LINE}px` }}
          onFocus={() => setFocused(true)}
          onBlur={() => {
            setFocused(false);
            const commit = draft.current!.take(props.content);
            setText(draft.current!.text);
            if (commit !== null) props.onCommit(commit);
          }}
          onInput={(e) => {
            draft.current!.input((e.target as HTMLTextAreaElement).value, props.target);
            setText(draft.current!.text);
          }}
          onKeyDown={(e) => {
            // Keys that confirm an IME composition belong to the input method.
            if (e.isComposing) return;
            if (e.key === 'Enter' && (e.altKey || e.shiftKey)) {
              // ⌥↩ / ⇧↩ start a new line in the cell, as in the grid's own editor.
              e.preventDefault();
              insertLineBreak(e.currentTarget);
            } else if (e.key === 'Enter') {
              e.preventDefault();
              const commit = draft.current!.take(props.content);
              setText(draft.current!.text);
              if (commit !== null) props.onCommit(commit);
              props.onDone();
            } else if (e.key === 'Escape') {
              e.preventDefault();
              draft.current!.reset(props.content);
              setText(props.content);
              props.onDone();
            }
          }}
        />
      </div>
    </div>
  );
}
