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
 *
 * SSE dispatch requires a genuine blank *line* — two consecutive terminators
 * — not merely a string that happens to end in one. `String.split` always
 * appends a trailing `''` when the input ends in a delimiter, and that
 * artifact is byte-identical to the `''` a real blank line produces; a body
 * that ends in exactly one `\r\n`/`\r`/`\n` would otherwise look dispatched
 * when the connection simply died right after the server's last line write.
 * The array `split` produces tells the two apart on its own, without a
 * second, separately-fallible regex: a genuine blank line leaves *two*
 * trailing `''` entries (the blank line itself, then its own boundary
 * artifact), where a body that merely ends in one terminator leaves only
 * one. (A regex re-checking the same thing, e.g. `/(?:\r\n|\r|\n){2}$/`,
 * looks equivalent but is not: to satisfy an exact `{2}`, the engine may
 * backtrack a single `\r\n` into a lone `\r` match plus a lone `\n` match,
 * which reports a false "doubled terminator" for exactly the single-CRLF
 * case this exists to catch — `split` never makes that substitution because
 * it always prefers the `\r\n` alternative and never needs to backtrack to
 * satisfy a rep count.) So when there's only the one artifact, it's dropped
 * before scanning, so the last (unterminated) frame's data survives in
 * `current` to the throw below instead of being mistaken for a dispatch.
 */
function extractSseData(body: string): string | null {
  const lines = body.split(/\r\n|\r|\n/);
  const last = lines.length - 1;
  if (lines[last] === '' && lines[last - 1] !== '') {
    lines.pop();
  }
  let current: string[] = [];
  for (const line of lines) {
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

function asRecord(value: unknown): Record<string, unknown> | null {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

/**
 * Turn one JSON-RPC message into a `ConvertResult`.
 *
 * `isError` is absent on success — mcp serialises with `exclude_unset`, so
 * "absent or false" is the success test and only an explicit `true` is a
 * failure. Error mapping arrives in the next task; for now anything that is
 * not a readable markdown result is a `proxy_error`, which fails loud rather
 * than handing the model a blank document.
 */
export function interpretJsonRpcMessage(message: unknown, status: number): ConvertResult {
  const envelope = asRecord(message);
  const result = asRecord(envelope?.result);
  if (!result) {
    return {
      ok: false,
      code: 'proxy_error',
      message: `the proxy answered ${status} with no JSON-RPC result`,
    };
  }
  const structured = asRecord(result.structuredContent);
  const markdown = structured?.markdown;
  if (typeof markdown !== 'string') {
    return {
      ok: false,
      code: 'proxy_error',
      message: `the proxy answered ${status} with a result carrying no markdown field`,
    };
  }
  return { ok: true, markdown };
}

export class AttachmentProxyClient {
  private readonly url: string;
  private readonly timeoutMs: number;
  private readonly fetchImpl: typeof fetch;

  constructor(opts: AttachmentProxyOptions) {
    this.url = opts.url;
    this.timeoutMs = opts.timeoutMs ?? DEFAULT_PROXY_TIMEOUT_MS;
    // Bound to globalThis rather than captured bare: an unbound `fetch`
    // reference throws "Illegal invocation" on some runtimes.
    this.fetchImpl = opts.fetchImpl ?? ((input, init) => globalThis.fetch(input, init));
  }

  private buildHeaders(): Record<string, string> {
    return {
      'Content-Type': 'application/json',
      // Both, and not negotiable: a Streamable HTTP endpoint answers 406 to a
      // POST that does not accept the streaming form, because it chooses the
      // framing per response.
      Accept: 'application/json, text/event-stream',
    };
  }

  async convertToMarkdown(req: ConvertRequest): Promise<ConvertResult> {
    const body = JSON.stringify({
      jsonrpc: '2.0',
      // A fixed id is correct here and not laziness: one request per
      // connection, no session, nothing to correlate against.
      id: 1,
      method: 'tools/call',
      params: { name: PROXY_TOOL_NAME, arguments: toWireArguments(req) },
    });

    const response = await this.fetchImpl(this.url, {
      method: 'POST',
      headers: this.buildHeaders(),
      body,
    });
    const contentType = response.headers.get('content-type');
    const text = await response.text();
    return interpretJsonRpcMessage(decodeJsonRpcBody(contentType, text), response.status);
  }
}
