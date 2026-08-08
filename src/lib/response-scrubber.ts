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
 * **Why a shape rule and not a field-name list.** Three distinct paths to the
 * same context blowout were found in eight days (`download-bytes`,
 * `get-mail-message-mime`, `expand: ["attachments"]`). Matching only the field
 * name `contentBytes` would be the same story-shaped guard one level down: it
 * covers the paths someone already thought of. The rule is therefore name
 * **plus** shape --
 *
 *   1. any field named `contentBytes` (the known Graph name), **or**
 *   2. any string over `BASE64_STRIP_THRESHOLD` characters that is valid
 *      base64.
 *
 * Rule 2 is what makes this a class guard: a tool upstream adds tomorrow that
 * returns bytes under a name nobody here has seen cannot regress this
 * deployment.
 *
 * The 4,096 floor is chosen against measured data. Graph ids in this deployment
 * are ~152 characters, so identifiers are never touched, and no real document
 * is smaller than 4 KB of base64.
 *
 * **Every strip is reported.** That reporting is the point of the module as
 * much as the stripping is: a fourth leak path should be discovered in a log
 * line naming the field, not in a context blowout.
 */

export const BASE64_STRIP_THRESHOLD = 4096;

/** Field names that are byte payloads whatever their length or shape. */
const BYTE_FIELD_NAMES = new Set(['contentBytes']);

/**
 * How deep the walk goes before it stops descending and replaces the subtree.
 *
 * Two jobs in one mechanism. Real Graph payloads nest ~10 deep, so 64 is far
 * past anything legitimate; hostile or malformed input that nests deeper is
 * replaced rather than recursed into, which keeps the walk off the stack limit.
 * A cyclic object -- `a.self = a` -- terminates here too, so cycles need no
 * separate bookkeeping. Both cases fail *closed*: the subtree is replaced by a
 * marker and reported, never passed through unexamined.
 */
const MAX_DEPTH = 64;

/**
 * Standard base64, whole string. Not base64url: Graph's `contentBytes` is
 * standard, and `-`/`_` are what most identifiers and tokens in these payloads
 * are built from, so accepting them would widen rule 2 towards exactly the
 * values it must not touch. A base64url value in a field *named* `contentBytes`
 * is still stripped, by rule 1.
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
  /** Decoded size of the payload in bytes, or 0 for a depth-capped subtree. */
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

function depthMarker(): string {
  return `<stripped: nesting deeper than ${MAX_DEPTH} levels, use read-document>`;
}

/**
 * The base64 candidate inside `text`, or null if it is not base64.
 *
 * Newlines are tolerated because line-wrapped base64 is normal in MIME, and a
 * payload would otherwise escape rule 2 by being 76 characters to a line.
 * *Only* `\r` and `\n`: stripping spaces too would join the words of ordinary
 * prose into a run of letters that can accidentally satisfy base64, which is
 * the false positive this module can least afford.
 */
function base64Candidate(text: string): string | null {
  if (isBase64(text)) return text;
  if (!/[\r\n]/.test(text)) return null;
  const joined = text.replace(/[\r\n]/g, '');
  return isBase64(joined) ? joined : null;
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
  if (BYTE_FIELD_NAMES.has(field)) {
    const named = base64Candidate(text);
    return named === null ? Buffer.byteLength(text, 'utf8') : decodedByteLength(named);
  }
  // Length first: every short string in every response reaches this line, and
  // most of them are never worth a regex.
  if (text.length <= BASE64_STRIP_THRESHOLD) return null;
  const candidate = base64Candidate(text);
  return candidate === null ? null : decodedByteLength(candidate);
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null) return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function scrub(
  node: unknown,
  path: string,
  field: string,
  stripped: StrippedField[],
  depth: number
): unknown {
  if (depth > MAX_DEPTH) {
    stripped.push({ path, field, bytes: 0 });
    return depthMarker();
  }

  if (typeof node === 'string') {
    const bytes = payloadBytes(field, node);
    if (bytes === null) return node;
    stripped.push({ path, field, bytes });
    return byteMarker(bytes);
  }

  if (Array.isArray(node)) {
    let changed = false;
    const out = node.map((item, index) => {
      const next = scrub(item, `${path}[${index}]`, String(index), stripped, depth + 1);
      if (next !== item) changed = true;
      return next;
    });
    return changed ? out : node;
  }

  if (isPlainObject(node)) {
    let changed = false;
    const out: Record<string, unknown> = {};
    for (const [key, item] of Object.entries(node)) {
      const next = scrub(item, `${path}.${key}`, key, stripped, depth + 1);
      if (next !== item) changed = true;
      out[key] = next;
    }
    return changed ? out : node;
  }

  // Numbers, booleans, null, undefined, and anything that is not a plain
  // object or array (a Date, a class instance) are returned as they came.
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
  const scrubbed = scrub(value, '$', '$', stripped, 0);
  return { value: scrubbed, stripped };
}
