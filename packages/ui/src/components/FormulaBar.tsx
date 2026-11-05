import { useEffect, useRef, useState } from 'preact/hooks';

/** Name box (address / go to) + the active cell's content. Commits on Enter, reverts on Escape. */
export function FormulaBar(props: {
  readonly address: string;
  readonly content: string;
  readonly onCommit: (text: string) => void;
  readonly onGoTo: (reference: string) => boolean;
  readonly onDone: () => void;
}) {
  const [text, setText] = useState(props.content);
  const [name, setName] = useState(props.address);
  const editing = useRef(false);
  const nameEditing = useRef(false);

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
        value={name}
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
      <input
        class="cs-formula-input"
        aria-label="Cell contents"
        spellcheck={false}
        value={text}
        onFocus={() => {
          editing.current = true;
        }}
        onBlur={() => {
          if (editing.current && text !== props.content) props.onCommit(text);
          editing.current = false;
        }}
        onInput={(e) => setText((e.target as HTMLInputElement).value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
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
  );
}
