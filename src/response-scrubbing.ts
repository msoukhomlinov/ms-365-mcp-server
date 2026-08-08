/**
 * The choke point that makes "no tool returns raw bytes" a property of the
 * server rather than of a list of tools.
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
 * `content` array of text blocks, an image/audio/resource block, or `content`
 * itself given as something other than the conventional array (a bare string,
 * a single object). Every one of those shapes is either walked by name or,
 * when unrecognised, passed through `scrubByteFields` whole rather than
 * skipped -- the same "an unrecognised shape is replaced, not passed through"
 * rule the scrubber's own depth cap applies one level down. It does not close
 * every byte-leak path in this deployment:
 *   - `graph-batch` accepts arbitrary sub-requests and can smuggle a GET
 *     against a suppressed path; that is a general bypass tracked separately,
 *     not something this wrapper can see into.
 *   - RFC 5322 MIME content (`get-mail-message-mime`-style) is text/plain, not
 *     base64, so this rule cannot match it -- that tool is suppressed instead.
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
 * read-document returns markdown the proxy produced. Scrubbing it would corrupt
 * a legitimate result -- a technical PDF with a base64 block inside a code fence
 * is enough -- to protect against bytes it structurally cannot contain.
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
 */
function describeStrippedField(field: StrippedField): string {
  const size = field.bytes > 0 ? `${field.bytes} bytes` : 'depth-capped, size unknown';
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
        'on this MCP server. Refusing to continue -- --attachment-proxy promises that no tool ' +
        'returns raw bytes, and an uninstalled scrubber cannot keep that promise silently.'
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
        // `content` was never a valid array to begin with, so there is no
        // conforming shape left to preserve it in. MCP requires `content` to
        // be an array of blocks regardless of what this module does, so
        // leaving the scrubbed value in whatever shape it already had (a bare
        // string, an object) would still be rejected downstream -- just with
        // the bytes now gone, trading a leak for an unconditional failure that
        // has nothing to do with bytes. Wrap it in a single text block, same
        // reasoning as the per-item case above: the one shape with nowhere
        // left for the marker to violate, and the one shape the transport can
        // always deliver.
        result.content = [textBlock(stringifyScrubbed(scrubbed.value))];
      }
    }

    if (result.structuredContent !== undefined) {
      const scrubbed = scrubByteFields(result.structuredContent);
      if (scrubbed.stripped.length > 0) {
        stripped.push(...scrubbed.stripped);
        result.structuredContent = scrubbed.value;
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
