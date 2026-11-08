/**
 * The one JSON spelling every `--json` reply uses: two-space indent, and a property whose value is
 * null or undefined is left out (so a plain number cell is three keys). Array elements are kept.
 */
export function toJson(payload: unknown): string {
  return JSON.stringify(
    payload,
    function (this: unknown, _key: string, value: unknown) {
      if ((value === null || value === undefined) && !Array.isArray(this)) return undefined;
      return value;
    },
    2,
  );
}

/** Single-line JSON (for diff lines). Same null rule. */
export function toCompactJson(payload: unknown): string {
  return JSON.stringify(payload, function (this: unknown, _key: string, value: unknown) {
    if ((value === null || value === undefined) && !Array.isArray(this)) return undefined;
    return value;
  });
}
