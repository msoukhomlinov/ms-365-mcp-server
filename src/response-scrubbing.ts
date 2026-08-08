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
 * not. It does not close every byte-leak path in this deployment:
 *   - `graph-batch` accepts arbitrary sub-requests and can smuggle a GET
 *     against a suppressed path; that is a general bypass tracked separately,
 *     not something this wrapper can see into.
 *   - RFC 5322 MIME content (`get-mail-message-mime`-style) is text/plain, not
 *     base64, so this rule cannot match it -- that tool is suppressed instead.
 *   - A minted ticket URL is neither `contentBytes` nor long base64, so this
 *     rule cannot match it either; `redactAttachmentSecrets` handles that.
 * Matching the disclosure already made at `isProxySuppressedGraphTool`: two
 * guards in this codebase silently disagreeing about what they cover would be
 * worse than either being honest alone.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { CallToolRequestSchema } from '@modelcontextprotocol/sdk/types.js';
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
    const result = (await original(request, extra)) as {
      content?: Array<{ type?: string; text?: string }>;
      structuredContent?: unknown;
    };
    const toolName = (request as { params?: { name?: string } }).params?.name ?? '(unknown)';
    if (SCRUBBER_BYPASS_TOOLS.has(toolName)) return result;

    const stripped: StrippedField[] = [];

    for (const item of result?.content ?? []) {
      if (item?.type !== 'text' || typeof item.text !== 'string') continue;
      const { value, wrapped } = parseToolText(item.text);
      const scrubbed = scrubByteFields(value);
      if (scrubbed.stripped.length === 0) continue;
      stripped.push(...scrubbed.stripped);
      item.text = wrapped
        ? String((scrubbed.value as { text?: unknown }).text ?? '')
        : JSON.stringify(scrubbed.value);
    }

    if (result?.structuredContent !== undefined) {
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
