/**
 * Strip byte payloads out of any tool result, at any depth, before it reaches
 * the model.
 *
 * This module carries the proxy-mode invariant -- *no tool on this server
 * returns raw bytes to the model* -- and it deliberately depends on nothing.
 * No server, no ticket store, no network: the unit that holds the invariant has
 * to be testable in isolation, and anything it imported would become a way for
 * the invariant to fail for reasons unrelated to bytes.
 *
 * **Every strip is reported.** That reporting is the point of the module as
 * much as the stripping is: a fourth leak path should be discovered in a log
 * line naming the field, not in a context blowout.
 */

/**
 * Minimum length, in characters, at which an unrecognised string is tested for
 * base64 shape. Strictly greater than: a 4,096-character string is kept.
 */
export const BASE64_STRIP_THRESHOLD = 4096;

/** Field names that are byte payloads whatever their length or shape. */
const BYTE_FIELD_NAMES = new Set(['contentBytes']);

/**
 * Standard base64, whole string. Not base64url: Graph's `contentBytes` is
 * standard, and `-`/`_` are what most identifiers and tokens in these payloads
 * are built from.
 *
 * Linear: the character class excludes `=`, so the greedy run cannot backtrack
 * into the padding.
 */
const BASE64_PATTERN = /^[A-Za-z0-9+/]*={0,2}$/;

/** One stripped payload, as reported to the caller (and thence to the log). */
export interface StrippedField {
  /** Full path of the stripped value, e.g. `$.value[0].attachments[1].contentBytes`. */
  path: string;
  /** Just the key, e.g. `contentBytes` -- the part worth grepping logs for. */
  field: string;
  /** Decoded size of the payload in bytes. */
  bytes: number;
}

export interface ScrubResult {
  /** The scrubbed value. The input is never mutated. */
  value: unknown;
  /** Every strip, in walk order. Empty means nothing was stripped. */
  stripped: StrippedField[];
}

/** What the model sees in place of the bytes. Names the tool that replaces it. */
function byteMarker(bytes: number): string {
  return `<stripped: ${bytes} bytes, use read-document>`;
}

function isBase64(text: string): boolean {
  if (text.length === 0 || text.length % 4 !== 0) return false;
  return BASE64_PATTERN.test(text);
}

/**
 * Bytes a base64 string decodes to, derived from its length rather than by
 * decoding it. Decoding a 291 KB payload just to measure it would allocate the
 * very buffer this module exists to keep out of the process.
 */
function decodedByteLength(base64: string): number {
  const padding = base64.endsWith('==') ? 2 : base64.endsWith('=') ? 1 : 0;
  return (base64.length / 4) * 3 - padding;
}

/**
 * Byte count if this string is a payload, or null if it should be left alone.
 *
 * A `contentBytes` that is not valid base64 is still stripped -- the name is
 * enough -- and measured as UTF-8, because its size is what the caller needs to
 * report and no decoded size exists.
 */
function payloadBytes(field: string, text: string): number | null {
  if (!BYTE_FIELD_NAMES.has(field)) return null;
  return isBase64(text) ? decodedByteLength(text) : Buffer.byteLength(text, 'utf8');
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null) return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function scrub(node: unknown, path: string, field: string, stripped: StrippedField[]): unknown {
  if (typeof node === 'string') {
    const bytes = payloadBytes(field, node);
    if (bytes === null) return node;
    stripped.push({ path, field, bytes });
    return byteMarker(bytes);
  }

  if (isPlainObject(node)) {
    let changed = false;
    const out: Record<string, unknown> = {};
    for (const [key, item] of Object.entries(node)) {
      const next = scrub(item, `${path}.${key}`, key, stripped);
      if (next !== item) changed = true;
      out[key] = next;
    }
    return changed ? out : node;
  }

  // Numbers, booleans, null, undefined, and anything that is not a plain
  // object (a Date, a class instance) are returned as they came.
  return node;
}

/**
 * Walk `value`, replacing every byte payload with a marker naming
 * `read-document`, and report what was replaced.
 *
 * Non-destructive: `value` is never mutated, and any branch containing no
 * payload is returned by reference rather than cloned, so a clean result is
 * `===` its input.
 */
export function scrubByteFields(value: unknown): ScrubResult {
  const stripped: StrippedField[] = [];
  const scrubbed = scrub(value, '$', '$', stripped);
  return { value: scrubbed, stripped };
}
