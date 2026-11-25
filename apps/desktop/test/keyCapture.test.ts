import { describe, expect, it } from 'vitest';
import { capturedKey, type KeyInput, keyCaptureAction } from '../src/keyCapture.ts';

const input = (init: Partial<KeyInput> & { key: string }): KeyInput => ({
  type: 'keyDown',
  code: '',
  shift: false,
  control: false,
  alt: false,
  meta: false,
  ...init,
});

describe('key capture', () => {
  it('leaves keys alone when not capturing', () => {
    expect(keyCaptureAction(false, input({ key: 'q', meta: true }))).toBe('pass');
  });

  it('holds keys back from the menu bar while capturing, and passes key-ups', () => {
    expect(keyCaptureAction(true, input({ key: 'q', meta: true }))).toBe('forward');
    expect(keyCaptureAction(true, input({ key: 'q', meta: true, type: 'keyUp' }))).toBe('pass');
  });

  it('always lets Escape out and ends capturing', () => {
    expect(keyCaptureAction(true, input({ key: 'Escape' }))).toBe('release');
  });

  it('hands the page the key as a keyboard event would', () => {
    expect(capturedKey(input({ key: 'W', code: 'KeyW', meta: true, shift: true }))).toEqual({
      key: 'W',
      code: 'KeyW',
      shiftKey: true,
      ctrlKey: false,
      altKey: false,
      metaKey: true,
    });
  });
});
