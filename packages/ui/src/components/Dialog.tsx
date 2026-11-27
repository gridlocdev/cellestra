import type { ComponentChildren } from 'preact';
import { useLayoutEffect, useRef } from 'preact/hooks';
import { trapTab } from './focusTrap.ts';

/**
 * A modal dialog: Escape and the backdrop close it, Tab cycles inside it, focus moves in on open
 * and goes back to the element that had it on close.
 */
export function Dialog(props: {
  readonly title?: string;
  readonly onClose: () => void;
  readonly class?: string;
  readonly children: ComponentChildren;
  readonly actions?: ComponentChildren;
  readonly onSubmit?: () => void;
}) {
  const ref = useRef<HTMLFormElement>(null);
  // Before the next key event: keys typed right after the dialog opens must not reach what is behind it.
  useLayoutEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    // The control marked autofocus wins, then a list's current item, then the first control: a
    // selector list would pick whichever comes first in the page, such as a Don't Save button.
    const form = ref.current;
    const first =
      form?.querySelector<HTMLElement>('[autofocus]') ??
      form?.querySelector<HTMLElement>('[data-focused]') ??
      form?.querySelector<HTMLElement>('input, select, button');
    first?.focus();
    return () => {
      if (previous?.isConnected && previous !== document.body) previous.focus();
    };
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
            } else if (ref.current) {
              trapTab(ref.current, e);
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
