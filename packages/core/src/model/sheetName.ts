/**
 * Worksheet-name rules (PRD §4 "rename"): non-empty, ≤ 31 chars, none of `: \ / ? * [ ]`, and no
 * leading/trailing apostrophe. Uniqueness is the Workbook's job. Spaces are legal (and quoted in refs).
 */
export const SheetName = {
  maxLength: 31,
  forbiddenChars: [':', '\\', '/', '?', '*', '[', ']'] as readonly string[],

  /** The human-readable reason a name is invalid, or null when it is fine. */
  validationError(name: string | null | undefined): string | null {
    if (!name || !name.trim()) return 'A sheet name cannot be empty.';
    if (name.length > SheetName.maxLength) return `A sheet name cannot be longer than ${SheetName.maxLength} characters.`;
    for (const c of name) if (SheetName.forbiddenChars.includes(c)) return `A sheet name cannot contain '${c}'.`;
    if (name.startsWith("'") || name.endsWith("'")) return 'A sheet name cannot start or end with an apostrophe.';
    return null;
  },

  isValid(name: string | null | undefined): boolean {
    return SheetName.validationError(name) === null;
  },

  validate(name: string | null | undefined): void {
    const error = SheetName.validationError(name);
    if (error) throw new RangeError(error);
  },
} as const;
