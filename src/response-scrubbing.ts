/**
 * The choke point that applies the byte rule of `scrubByteFields` to every
 * `tools/call` result on its way out, instead of tool by tool. What that makes a
 * property of this wrapper is its *reach*: no tool has to opt in, and a tool
 * added tomorrow is covered without anyone editing a list. It does not make the
 * rule itself complete -- see **Coverage** below.
 *
 * This used to open by claiming it made "no tool returns raw bytes" a property
 * of the server. That guarantee was withdrawn in 3025efe (PR #19) as false; the
 * reasoning is recorded above `CLASSIFIED` in `test/absence-claims.test.ts`. The
 * Coverage note below was already the accurate account of the same file, so the
 * opening sentence was not merely incomplete -- it asserted what the rest of
 * this comment goes on to refute.
 *
 * Modelled on `normalize-tool-schema.ts`, which decorates `tools/list` the same
 * way, with one deliberate difference: that one falls back to a no-op and a
 * warning if the SDK internals move, because a missing $ref rewrite degrades
 * gracefully. This one throws. A server that believes it is scrubbing and is not
 * is worse than one that refuses to start, and the whole point of putting the
 * enforcement here rather than in each tool is that it cannot be partially on.
 *
 * **Coverage.** This closes "a tool result contains `contentBytes` or a long
 * base64 string" -- see `scrubByteFields` for exactly what that rule is and is
 * not -- regardless of where in the result that string lives: a conventional
 * `content` array of text blocks, an image/audio/resource block, `content`
 * itself given as something other than the conventional array (a bare string,
 * a single object), or `structuredContent`. Every one of those shapes is
 * either walked by name, or -- when unrecognised, or when scrubbing would
 * leave a shape its own schema cannot hold (`content` must be an array;
 * `structuredContent` must be a plain object; an image/audio/resource block's
 * `data`/`resource.blob` must be valid base64, which a human-readable marker
 * never is) -- converted into a text block in `content` rather than assigned
 * back broken. That conversion runs even when nothing was actually stripped
 * for `content`-as-a-whole and `structuredContent` (there is no legitimate
 * value either can hold outside their required shape, so fixing costs nothing
 * real); it does *not* run for an individual `content` array item with
 * nothing stripped, because a small, valid image block is real and must
 * survive untouched. It does not close every byte-leak path in this
 * deployment:
 *   - `graph-batch` accepts arbitrary sub-requests and can smuggle a GET
 *     against a suppressed path; that is a general bypass tracked separately,
 *     not something this wrapper can see into.
 *   - A non-JSON text body is not base64, so this rule cannot match it. It
 *     arrives under `rawResponse` rather than `contentBytes`, which rule 1 does
 *     not name either. `get-mail-message-mime` is the RFC 5322 case and is
 *     suppressed instead, but suppression is GET plus a path ending `/$value`,
 *     and `get-meeting-transcript-content` (`text/vtt`) and
 *     `get-onenote-page-content` end in `/content` -- so they stay registered in
 *     proxy mode with neither rule reaching their bodies.
 *   - A minted ticket URL is neither `contentBytes` nor long base64, so this
 *     rule cannot match it either; `redactAttachmentSecrets` handles that.
 *   - A tool result that is not an object at all -- no `content` or
 *     `structuredContent` field to find on a bare string, number, or `null` --
 *     is passed through unscrubbed rather than guessed at. No spec-conformant
 *     `CallToolResult` takes that shape, so this is a narrow boundary, not a
 *     realistic leak: there is nowhere for a byte payload to be *named* in a
 *     result with no fields.
 * Matching the disclosure already made at `isProxySuppressedGraphTool`: two
 * guards in this codebase silently disagreeing about what they cover would be
 * worse than either being honest alone.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { CallToolRequestSchema, type ServerResult } from '@modelcontextprotocol/sdk/types.js';
import logger from './logger.js';
import { scrubByteFields, type StrippedField } from './lib/response-scrubber.js';

/**
 * Tools whose results are never scrubbed.
 *
 * read-document returns markdown the proxy produced, never bytes -- so this is
 * defense-in-depth, not a fix for something that happens routinely. A base64
 * block sitting inside a markdown code fence does NOT actually trip rule 2 on
 * its own (the heading and fence markers are not base64 characters, and rule 2
 * requires the *entire* field to match); what this guards against is the rarer
 * case where read-document's markdown coincidentally contains some run of text
 * that, on its own, satisfies the shape rule. Scrubbing that would corrupt a
 * legitimate result to protect against bytes read-document structurally cannot
 * contain in the first place.
 */
const SCRUBBER_BYPASS_TOOLS = new Set(['read-document']);

/**
 * A tool result's text is usually a JSON document and sometimes it is not.
 * The scrubber's second rule is about *fields*, so a bare string body would slip
 * past it entirely; wrapping a non-JSON body in a one-field object puts it under
 * the same rule instead of leaving a hole shaped exactly like the payload with
 * no JSON envelope to hide in.
 */
function parseToolText(text: string): { value: unknown; wrapped: boolean } {
  try {
    return { value: JSON.parse(text), wrapped: false };
  } catch {
    return { value: { text }, wrapped: true };
  }
}

/** True for a conventional `{ type: 'text', text: string }` content block. */
function isTextContentItem(item: unknown): item is { type: 'text'; text: string } {
  return (
    typeof item === 'object' &&
    item !== null &&
    (item as Record<string, unknown>).type === 'text' &&
    typeof (item as Record<string, unknown>).text === 'string'
  );
}

/**
 * A text content block, the one MCP content-block shape with no format
 * constraint on its payload. Used to hold a scrubbed value that has nowhere
 * else safe to live -- see `scrubContentItem` and the top-level `content`
 * fallback in `installResponseScrubbing` for why that matters.
 */
function textBlock(text: string): { type: 'text'; text: string } {
  return { type: 'text', text };
}

/** `value` as a string, JSON-encoding it first if it is not one already. */
function stringifyScrubbed(value: unknown): string {
  return typeof value === 'string' ? value : JSON.stringify(value);
}

/** True for `{}`-shaped values -- not an array, not `null`, not a primitive. */
function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/**
 * Scrub one entry of a `content` array.
 *
 * A conventional text block gets the JSON-aware treatment: its `text` is parsed,
 * scrubbed as a value, and re-serialized, so a `contentBytes` field three levels
 * into a JSON body is still caught.
 *
 * Anything else -- an image or audio block whose bytes live in a `data` field,
 * a resource block whose bytes live in `resource.blob`, or a shape this module
 * has never seen -- has no dedicated parsing rule, so it goes through
 * `scrubByteFields` directly rather than being skipped: the shape rule was
 * never text-specific (rule 2 is "any long base64 string, wherever it is").
 * But the *replacement* value that rule produces is a human-readable marker,
 * not valid base64 -- and an image/audio block's `data` (or a resource block's
 * `resource.blob`) is contractually required by its own content-block schema
 * to decode as base64. Handing back a block that still claims to be `image`
 * but no longer satisfies ImageContentSchema would make the whole tool result
 * invalid at the transport layer, for a reason that has nothing to do with the
 * bytes this module exists to catch -- trading a possible byte leak for a
 * guaranteed broken tool call. So a scrubbed non-text block is converted
 * wholesale into a text block instead: same information, a shape with nowhere
 * left for the marker to violate.
 */
function scrubContentItem(item: unknown, stripped: StrippedField[]): unknown {
  if (isTextContentItem(item)) {
    const { value, wrapped } = parseToolText(item.text);
    const scrubbed = scrubByteFields(value);
    if (scrubbed.stripped.length === 0) return item;
    stripped.push(...scrubbed.stripped);
    item.text = wrapped
      ? String((scrubbed.value as { text?: unknown }).text ?? '')
      : JSON.stringify(scrubbed.value);
    return item;
  }
  const scrubbed = scrubByteFields(item);
  if (scrubbed.stripped.length === 0) return item;
  stripped.push(...scrubbed.stripped);
  return textBlock(stringifyScrubbed(scrubbed.value));
}

/** Sentinel `bytes` value this module (not the scrubber) uses on a synthetic
 * `StrippedField` it manufactures for a shape-only rescue -- see the
 * `content`/`structuredContent` handling below. Negative and distinct from
 * the scrubber's own `bytes: 0` depth-cap convention, so the two never
 * describe themselves the same way in the log (see `describeStrippedField`).
 */
const SHAPE_RESCUE_BYTES = -1;

/**
 * Render one stripped field for the warn line.
 *
 * A depth-capped subtree reports `bytes: 0` by design -- the walk refused to
 * look inside it, so there is no size to report, and 0 is a *report* of "size
 * unknown, subtree replaced" rather than a claim that nothing was there. It is
 * called out explicitly here (rather than folded into a silent 0) so a reader
 * of the log line -- or of a total computed from it -- sees "depth-capped" and
 * not a size. If this ever gets summed elsewhere, that sum will understate for
 * every depth-capped entry; the fix there is to skip or flag those entries by
 * name, not to have this module fabricate a size it doesn't have.
 *
 * A *shape-only* rescue (`content` or `structuredContent` fixed into a valid
 * shape with no bytes anywhere in it) is a third case this function has to
 * tell apart from both: it is not a byte count, and calling it "depth-capped"
 * would blame the wrong mechanism for the rescue. `SHAPE_RESCUE_BYTES` (a
 * negative sentinel this module manufactures itself, never something
 * `scrubByteFields` produces) keeps that case worded honestly too.
 */
function describeStrippedField(field: StrippedField): string {
  const size =
    field.bytes > 0
      ? `${field.bytes} bytes`
      : field.bytes === SHAPE_RESCUE_BYTES
        ? 'shape rescue, no bytes involved'
        : 'depth-capped, size unknown';
  return `${field.path || '(root)'}.${field.field} (${size})`;
}

export function installResponseScrubbing(server: McpServer): void {
  const lowLevel = server.server;
  const handlers = (
    lowLevel as unknown as {
      _requestHandlers?: Map<string, (request: unknown, extra: unknown) => Promise<unknown>>;
    }
  )._requestHandlers;
  const original = handlers?.get('tools/call');
  if (!original) {
    throw new Error(
      'Cannot install the attachment-proxy response scrubber: no tools/call handler is registered ' +
        'on this MCP server. Refusing to continue -- --attachment-proxy undertakes to strip byte ' +
        'payloads from tool results, and an uninstalled scrubber cannot keep that undertaking silently.'
    );
  }

  lowLevel.setRequestHandler(CallToolRequestSchema, async (request, extra) => {
    const raw = await original(request, extra);
    const toolName = (request as { params?: { name?: string } }).params?.name ?? '(unknown)';
    // The two early returns below hand back `raw` exactly as `original` produced
    // it -- untouched, whatever its shape -- so the cast is just telling the
    // compiler the value is going out the same door it came in, not asserting
    // anything about its shape.
    if (SCRUBBER_BYPASS_TOOLS.has(toolName)) return raw as ServerResult;

    // A conformant result is always an object; anything else (missing entirely,
    // or a bare primitive some hand-rolled handler returned) has no `content` or
    // `structuredContent` field to scrub and is returned as-is rather than
    // crashing the request trying to read a property off it.
    if (raw === null || typeof raw !== 'object') return raw as ServerResult;

    const result = raw as { content?: unknown; structuredContent?: unknown };
    const stripped: StrippedField[] = [];

    if (Array.isArray(result.content)) {
      const content: unknown[] = result.content;
      for (let i = 0; i < content.length; i++) {
        content[i] = scrubContentItem(content[i], stripped);
      }
    } else if (result.content !== undefined) {
      // `content` is present but not the conventional array of blocks -- a bare
      // string, a single object, or some other shape. A loop that only knows
      // how to walk arrays either silently does nothing here (a string iterates
      // as characters, none of which look like a content block) or throws (a
      // plain object isn't iterable at all) -- both are a hole in the choke
      // point. Unknown top-level shape gets MORE scrutiny instead: run the
      // whole value through scrubByteFields directly, the same "an
      // unrecognised shape is replaced, not passed through" rule the
      // scrubber's own depth cap already applies one level down.
      const scrubbed = scrubByteFields(result.content);
      if (scrubbed.stripped.length > 0) {
        stripped.push(...scrubbed.stripped);
      }
      // Being inside this branch at all already means `content` is not an
      // array, and content's schema always requires one -- unlike a content
      // ARRAY ITEM, where a small legitimate image block is real and must
      // survive untouched (see scrubContentItem's own gate), there is no
      // legitimate value this branch can hold, so fixing it costs nothing
      // real. Record the rescue even when no bytes were involved
      // (`SHAPE_RESCUE_BYTES`) so a shape-only fix stays visible in the log,
      // and wrap the (possibly byte-scrubbed) value in a single text block:
      // nowhere for a marker to violate, and the one shape the transport can
      // always deliver.
      if (scrubbed.stripped.length === 0) {
        stripped.push({ path: '$', field: 'content', bytes: SHAPE_RESCUE_BYTES });
      }
      result.content = [textBlock(stringifyScrubbed(scrubbed.value))];
    }

    if (result.structuredContent !== undefined) {
      const scrubbed = scrubByteFields(result.structuredContent);
      if (scrubbed.stripped.length > 0) {
        stripped.push(...scrubbed.stripped);
      }
      if (isPlainObject(scrubbed.value)) {
        result.structuredContent = scrubbed.value;
      } else {
        // structuredContent's own schema (a record: `{[key: string]: unknown}`)
        // requires a plain object, the same way content's schema requires an
        // array -- and it can end up here two ways. Either scrubByteFields
        // replaced the WHOLE value with a marker string because
        // structuredContent itself carried the bytes (already reported above,
        // in `stripped`), or structuredContent was never an object to begin
        // with and had nothing to strip at all (nothing above to report,
        // because there were no bytes -- but the shape is still wrong).
        // Assigning either back would violate the schema for a reason that has
        // nothing to do with the bytes this module exists to catch -- fails
        // closed for `content` and fails open for `structuredContent` is not a
        // choke point. Record the rescue even when no bytes were involved
        // (`SHAPE_RESCUE_BYTES`) so a shape-only fix is still visible in the
        // log rather than silent, drop the field -- nothing downstream needs
        // an invalid `structuredContent` more than it needs a missing one --
        // and give the marker the one home that's always valid: a text block
        // in `content`.
        if (scrubbed.stripped.length === 0) {
          stripped.push({ path: '$', field: 'structuredContent', bytes: SHAPE_RESCUE_BYTES });
        }
        delete result.structuredContent;
        const marker = textBlock(stringifyScrubbed(scrubbed.value));
        if (Array.isArray(result.content)) {
          result.content.push(marker);
        } else {
          result.content = [marker];
        }
      }
    }

    if (stripped.length > 0) {
      // Naming the field is the whole reporting requirement: a fourth byte path
      // should be discovered in a log line, not in a context blowout.
      logger.warn(
        `Response scrubber stripped ${stripped.length} byte field(s) from ${toolName}: ` +
          stripped.map(describeStrippedField).join(', ') +
          '. Use read-document to read this content as markdown.'
      );
    }

    return result;
  });
}
