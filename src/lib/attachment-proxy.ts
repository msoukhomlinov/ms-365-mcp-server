import logger from '../logger.js';

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

/**
 * Environment variable holding the proxy's bearer token, when it wants one.
 *
 * Read here rather than passed in because `AttachmentProxyOptions` is a fixed
 * contract this section shares with its siblings, and a token is not a
 * behaviour the caller chooses -- it is a fact about the endpoint. Read per
 * call rather than in the constructor so a rotated token takes effect without
 * a restart, and so a test can set it without rebuilding the client.
 */
export const PROXY_TOKEN_ENV = 'MS365_MCP_ATTACHMENT_PROXY_TOKEN';

/**
 * The proxy-origin codes this server promises its own callers, from the spec's
 * error table. Everything else the proxy says is real but is not ours to
 * promise — a generic contract cannot adopt one implementation's vocabulary —
 * so it arrives as `proxy_error` with the raw code preserved verbatim.
 */
export const CONTRACT_ERROR_CODES: ReadonlySet<string> = new Set([
  'unsupported_format',
  'too_large',
  'password_required',
  'conversion_failed',
  'fetch_failed',
]);

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
 * The coded payload out of the text channel, for a proxy that does not send
 * `structuredContent`. The contract says "an error object carrying a string
 * code", not "on this particular channel".
 */
function parseTextContent(result: Record<string, unknown>): Record<string, unknown> | null {
  const content = Array.isArray(result.content) ? result.content : [];
  for (const block of content) {
    const record = asRecord(block);
    if (!record || typeof record.text !== 'string') continue;
    try {
      const parsed = asRecord(JSON.parse(record.text));
      if (parsed) return parsed;
    } catch {
      continue;
    }
  }
  return null;
}

function mapProxyError(payload: Record<string, unknown> | null): ConvertResult {
  const rawMessage =
    typeof payload?.message === 'string'
      ? payload.message
      : 'the proxy reported an error with no message';
  const detail = asRecord(payload?.detail);
  const status =
    detail && typeof detail.status === 'number' ? ` (upstream status ${detail.status})` : '';
  const rawCode = typeof payload?.code === 'string' ? payload.code : '';

  if (rawCode === '') {
    return {
      ok: false,
      code: 'proxy_error',
      message: `the proxy reported an error with no code: ${rawMessage}${status}`,
    };
  }
  if (CONTRACT_ERROR_CODES.has(rawCode)) {
    return { ok: false, code: rawCode, message: `${rawMessage}${status}` };
  }
  // Verbatim, and first: the raw code has to survive a log truncation and a
  // grep, which it does not if it is buried mid-sentence.
  return { ok: false, code: 'proxy_error', message: `${rawCode}: ${rawMessage}${status}` };
}

/**
 * Turn one JSON-RPC message into a `ConvertResult`.
 *
 * `isError` is absent on success — mcp serialises with `exclude_unset` — so
 * "absent" and "false" both read as success and only an explicit `true`
 * triggers error mapping; `result.isError === true` is that exact test, not
 * an inference from the shape of `structuredContent`.
 */
export function interpretJsonRpcMessage(message: unknown, status: number): ConvertResult {
  const envelope = asRecord(message);
  const rpcError = asRecord(envelope?.error);
  if (rpcError) {
    // Transport- or dispatcher-level refusal: auth, an unserved method, a
    // malformed envelope. The proxy answered, so it is not `proxy_unreachable`.
    return {
      ok: false,
      code: 'proxy_error',
      message: `the proxy refused the call (HTTP ${status}, JSON-RPC ${String(
        rpcError.code
      )}): ${String(rpcError.message ?? '')}`,
    };
  }

  const result = asRecord(envelope?.result);
  if (!result) {
    return {
      ok: false,
      code: 'proxy_error',
      message: `the proxy answered ${status} with no JSON-RPC result`,
    };
  }

  const payload = asRecord(result.structuredContent) ?? parseTextContent(result);
  if (result.isError === true) return mapProxyError(payload);

  const markdown = payload?.markdown;
  if (typeof markdown !== 'string') {
    return {
      ok: false,
      code: 'proxy_error',
      message: `the proxy answered ${status} with a result carrying no markdown field`,
    };
  }
  return { ok: true, markdown };
}

/**
 * The most specific sentence available about why a fetch failed.
 *
 * Node's fetch reports every transport failure as a bare `TypeError: fetch
 * failed` and puts the real reason -- ECONNREFUSED, EAI_AGAIN, a TLS error --
 * on `cause`. Reporting only the outer message tells an operator nothing about
 * whether the proxy is down, misspelled, or unresolvable.
 */
function describeTransportFailure(error: unknown): string {
  const base = error instanceof Error ? error.message : String(error);
  const cause = error instanceof Error ? (error as { cause?: unknown }).cause : undefined;
  return cause instanceof Error ? `${base}: ${cause.message}` : base;
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
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      // Both, and not negotiable: a Streamable HTTP endpoint answers 406 to a
      // POST that does not accept the streaming form, because it chooses the
      // framing per response.
      Accept: 'application/json, text/event-stream',
    };
    const token = process.env[PROXY_TOKEN_ENV];
    // Empty-string-is-absent on purpose: an unset compose variable interpolates
    // to "", and `Bearer ` is a credential shape a server may log as one.
    if (token) headers.Authorization = `Bearer ${token}`;
    return headers;
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

    const startedAt = Date.now();
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.timeoutMs);

    let status: number;
    let contentType: string | null;
    let text: string;
    try {
      const response = await this.fetchImpl(this.url, {
        method: 'POST',
        headers: this.buildHeaders(),
        body,
        signal: controller.signal,
      });
      status = response.status;
      contentType = response.headers.get('content-type');
      // Inside the guarded block on purpose. In SSE framing the body *is* the
      // conversion: headers arrive immediately and the data frame arrives when
      // the document is done, so a timeout that ended at the headers would
      // never fire on the failure it exists for.
      text = await response.text();
    } catch (error) {
      const elapsedMs = Date.now() - startedAt;
      const aborted = error instanceof Error && error.name === 'AbortError';
      const reason = aborted
        ? `no response within ${this.timeoutMs} ms`
        : describeTransportFailure(error);
      // Debug, not warn. This client does not retry and has no notion of an
      // "attempt" -- `read-document` in graph-tools.ts does, and it is the one
      // WARN an operator should see per call, worded with attempt context
      // ("attempt 1 of 2" vs "still unreachable after retry"). A warn at both
      // layers means a single wedged proxy produces two lines with two clocks
      // and two wordings per attempt (four across a retry) for what should be
      // one clear signal -- worse for the exact failure this exists to catch
      // (a service that ran 500s for 19 hours while its container reported
      // healthy) than a single line would be. Do NOT restore this to warn:
      // the caller owns that signal now. Kept at debug rather than dropped --
      // this is a standalone unit with its own tests, and its detail (the raw
      // transport reason before the caller's retry framing) is worth having
      // available without needing a WARN's severity to justify existing.
      logger.debug(`[ATTACHMENT PROXY] unreachable: ${this.url} after ${elapsedMs} ms — ${reason}`);
      return {
        ok: false,
        code: 'proxy_unreachable',
        message: `the document proxy at ${this.url} did not answer: ${reason}`,
      };
    } finally {
      clearTimeout(timer);
    }

    let message: unknown;
    try {
      message = decodeJsonRpcBody(contentType, text);
    } catch {
      // A gateway's HTML, a truncated stream, a proxy that answered in prose.
      // The body is deliberately not quoted: it is upstream text this server
      // has not audited, and it would land in the model's context.
      return {
        ok: false,
        code: 'proxy_error',
        message: `the proxy answered ${status} with a body this client could not read as JSON-RPC`,
      };
    }
    return interpretJsonRpcMessage(message, status);
  }
}
