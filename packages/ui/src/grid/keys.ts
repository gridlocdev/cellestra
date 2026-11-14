/** The parts of a keydown that tell a typed character from a composition. */
export interface KeyLike {
  readonly key: string;
  readonly keyCode: number;
  readonly isComposing: boolean;
}

/** keyCode browsers report for a key an IME (or a dead key, in some) is processing. */
const IME_KEY_CODE = 229;

/**
 * Whether a keydown on the grid starts a composed character: a dead key (´ then e for é) or an IME
 * (keyCode 229, `Process`). Its text only arrives later as input, and only in an editable element,
 * so the grid must open its editor now, without preventing the key, for the character to land.
 */
export function startsComposition(e: KeyLike): boolean {
  return e.key === 'Dead' || e.key === 'Process' || e.isComposing || e.keyCode === IME_KEY_CODE;
}
