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
