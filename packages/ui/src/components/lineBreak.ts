/**
 * Type a line break at the caret. `insertText` keeps it on the field's own undo stack (⌘Z takes
 * it back) and fires `input`; where that is unsupported it falls back to editing the value.
 */
export function insertLineBreak(field: HTMLTextAreaElement): void {
  if (document.execCommand?.('insertText', false, '\n')) return;
  field.setRangeText('\n', field.selectionStart, field.selectionEnd, 'end');
  field.dispatchEvent(new Event('input', { bubbles: true }));
}
