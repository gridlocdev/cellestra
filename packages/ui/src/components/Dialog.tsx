import type { ComponentChildren } from 'preact';
import { useEffect, useRef } from 'preact/hooks';

/** A modal dialog: Escape and the backdrop close it, focus moves in and is restored on close. */
export function Dialog(props: {
  readonly title?: string;
  readonly onClose: () => void;
  readonly class?: string;
  readonly children: ComponentChildren;
  readonly actions?: ComponentChildren;
  readonly onSubmit?: () => void;
}) {
  const ref = useRef<HTMLFormElement>(null);
  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    const first = ref.current?.querySelector<HTMLElement>('[autofocus], input, select, button');
    first?.focus();
    return () => previous?.focus();
  }, []);
  return (
    <div
      class="cs-backdrop"
      onPointerDown={(e) => {
        if (e.target === e.currentTarget) props.onClose();
      }}
    >
      <div class={`cs-dialog ${props.class ?? ''}`} role="dialog" aria-modal="true" aria-label={props.title}>
        <form
          ref={ref}
          style={{ display: 'contents' }}
          onKeyDown={(e) => {
            if (e.key === 'Escape') {
              e.preventDefault();
              e.stopPropagation();
              props.onClose();
            }
          }}
          onSubmit={(e) => {
            e.preventDefault();
            props.onSubmit?.();
          }}
        >
          {props.title && <h2>{props.title}</h2>}
          <div class="cs-body">{props.children}</div>
          {props.actions && <div class="cs-actions">{props.actions}</div>}
        </form>
      </div>
    </div>
  );
}
