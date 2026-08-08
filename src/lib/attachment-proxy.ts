/**
 * One stateless JSON-RPC call to a document-conversion proxy.
 *
 * The proxy contract is deliberately generic — "stateless MCP over Streamable
 * HTTP, exposing `convert_to_markdown(uri, pages?, offset?, max_chars?)`" —
 * and nothing in this file names docglean. What it *does* encode is the
 * transport: a single POST with no session handshake, an `Accept` header
 * carrying both media types, and a response that may arrive either as
 * `application/json` or framed as `text/event-stream`.
 *
 * That last point is the whole reason this module exists as its own unit with
 * its own tests. On 2026-08-07 a defect shipped because every test mocked the
 * layer *below* the wire and handed it a truth the real client could not tell,
 * so the derivation was never exercised and every attachment served as an
 * empty 200. Here the equivalent trap is mocking this client from the tool
 * layer and never running a real frame through it. Tests mock `fetch`, never
 * this class.
 */

/** Proxy timeout. A 195 KB PDF converted in 0.112 s live; this is a wedge detector. */
export const DEFAULT_PROXY_TIMEOUT_MS = 60_000;

/** The one tool the contract requires of any `--attachment-proxy` target. */
export const PROXY_TOOL_NAME = 'convert_to_markdown';

export interface AttachmentProxyOptions {
  /** Full endpoint URL including the MCP path, e.g. `http://docglean:8080/mcp`. */
  url: string;
  timeoutMs?: number;
  fetchImpl?: typeof fetch;
}

export interface ConvertRequest {
  uri: string;
  pages?: string;
  offset?: number;
  maxChars?: number;
}

export type ConvertResult =
  | { ok: true; markdown: string }
  | { ok: false; code: string; message: string };

/**
 * camelCase in, snake_case out.
 *
 * Optional parameters are omitted rather than sent as null: a null `max_chars`
 * fails the tool's own argument validation and comes back as a coded error,
 * where an absent one correctly means "use the server's default".
 */
export function toWireArguments(req: ConvertRequest): Record<string, unknown> {
  const args: Record<string, unknown> = { uri: req.uri };
  if (req.pages !== undefined) args.pages = req.pages;
  if (req.offset !== undefined) args.offset = req.offset;
  if (req.maxChars !== undefined) args.max_chars = req.maxChars;
  return args;
}

/**
 * The data payload of the first SSE event that carries one, or null if the
 * stream never carries any.
 *
 * Written against the framing rather than against one server's output: comment
 * frames (`: ping - …`) are skipped, `event:`/`id:`/`retry:` lines are ignored,
 * a frame whose data is empty (mcp's resumability priming event) is passed
 * over rather than parsed as JSON, and multiple `data:` lines in one frame are
 * joined with newlines as the SSE spec requires. CRLF, CR and LF all separate
 * lines, because sse_starlette defaults to CRLF and other servers do not.
 *
 * A stream that ends with data accumulated but no terminating blank line is
 * truncated, not merely quiet: under SSE the body IS the conversion, so a
 * frame that never closed means the conversion never finished. Returning that
 * partial accumulation as success would be a clean, well-formed, wrong answer
 * — the same failure class as the empty 200 of 2026-08-07 — so this throws
 * instead of degrading silently.
 */
function extractSseData(body: string): string | null {
  let current: string[] = [];
  for (const line of body.split(/\r\n|\r|\n/)) {
    if (line === '') {
      const joined = current.join('\n');
      current = [];
      if (joined !== '') return joined;
      continue;
    }
    if (line.startsWith(':') || !line.startsWith('data:')) continue;
    const value = line.slice('data:'.length);
    current.push(value.startsWith(' ') ? value.slice(1) : value);
  }
  if (current.length > 0) {
    throw new Error('the proxy event stream was truncated before its terminating blank line');
  }
  return null;
}

/**
 * Parse a Streamable HTTP response body into its JSON-RPC message.
 *
 * The transport picks the framing, not the caller: `text/event-stream` when the
 * server streams (the default, and what the live proxy does) and
 * `application/json` when it is in JSON-response mode. Handling only one of the
 * two is exactly the class of defect that served every attachment as an empty
 * 200 on 2026-08-07 — a well-formed response the client could not read.
 */
export function decodeJsonRpcBody(contentType: string | null, body: string): unknown {
  const isSse = (contentType ?? '').toLowerCase().includes('text/event-stream');
  const payload = isSse ? extractSseData(body) : body;
  if (payload === null) {
    throw new Error('the proxy sent an event stream with no data frame in it');
  }
  return JSON.parse(payload);
}
