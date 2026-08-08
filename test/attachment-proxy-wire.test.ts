import { describe, it, expect } from 'vitest';
import { toWireArguments, decodeJsonRpcBody } from '../src/lib/attachment-proxy.js';

/**
 * The TS interface is camelCase and the wire is snake_case. A `maxChars` that
 * never becomes `max_chars` is not an error anywhere: docglean applies its own
 * default cap and returns a perfectly good, wrongly-sized document.
 */
describe('toWireArguments', () => {
  it('renames maxChars to max_chars and leaves the rest alone', () => {
    expect(
      toWireArguments({ uri: 'https://p/a?t=1', pages: '1-3', offset: 40, maxChars: 20000 })
    ).toEqual({
      uri: 'https://p/a?t=1',
      pages: '1-3',
      offset: 40,
      max_chars: 20000,
    });
  });

  it('omits every optional parameter that was not given, rather than sending null', () => {
    const args = toWireArguments({ uri: 'https://p/a?t=1' });
    expect(args).toEqual({ uri: 'https://p/a?t=1' });
    // Absent, not present-and-undefined: JSON.stringify drops undefined, but
    // `in` is what a reviewer can check, and a null here is invalid_max_chars.
    expect('max_chars' in args).toBe(false);
    expect('offset' in args).toBe(false);
    expect('pages' in args).toBe(false);
  });

  it('keeps a zero offset, which is a real value and not an absent one', () => {
    expect(toWireArguments({ uri: 'u', offset: 0 })).toEqual({ uri: 'u', offset: 0 });
  });

  it('never invents an auth_profile', () => {
    expect('auth_profile' in toWireArguments({ uri: 'u' })).toBe(false);
  });
});

/**
 * Streamable HTTP answers a POST either as `application/json` or as an SSE
 * stream, and the live proxy answers with SSE (mcp's `streamable_http_app`
 * builds an `EventSourceResponse` unless `json_response=True`). sse_starlette
 * frames with CRLF, names the event `message`, and interleaves `: ping - …`
 * comment frames every 15 s — which a 60 s conversion will see.
 */
describe('decodeJsonRpcBody', () => {
  const message = { jsonrpc: '2.0', id: 1, result: { structuredContent: { markdown: '# hi' } } };

  it('reads the JSON-RPC message out of a text/event-stream frame', () => {
    const body = `event: message\r\ndata: ${JSON.stringify(message)}\r\n\r\n`;
    expect(decodeJsonRpcBody('text/event-stream; charset=utf-8', body)).toEqual(message);
  });

  it('reads the same message out of a plain application/json body', () => {
    expect(decodeJsonRpcBody('application/json', JSON.stringify(message))).toEqual(message);
  });

  it('skips the keep-alive comment frames a slow conversion interleaves', () => {
    const body =
      `: ping - 2026-08-08 07:00:00.000000+00:00\r\n\r\n` +
      `: ping - 2026-08-08 07:00:15.000000+00:00\r\n\r\n` +
      `event: message\r\ndata: ${JSON.stringify(message)}\r\n\r\n`;
    expect(decodeJsonRpcBody('text/event-stream', body)).toEqual(message);
  });

  it('skips a priming frame that carries an id and empty data', () => {
    const body =
      `id: 0\r\ndata: \r\n\r\n` + `event: message\r\ndata: ${JSON.stringify(message)}\r\n\r\n`;
    expect(decodeJsonRpcBody('text/event-stream', body)).toEqual(message);
  });

  it('joins a multi-line data payload the way the SSE spec says to', () => {
    const body = 'event: message\ndata: {"jsonrpc":"2.0",\ndata: "id":1}\n\n';
    expect(decodeJsonRpcBody('text/event-stream', body)).toEqual({ jsonrpc: '2.0', id: 1 });
  });

  it('accepts LF-only framing as well as CRLF', () => {
    const body = `event: message\ndata: ${JSON.stringify(message)}\n\n`;
    expect(decodeJsonRpcBody('text/event-stream', body)).toEqual(message);
  });

  it('throws rather than returning undefined when a stream carries no data', () => {
    expect(() => decodeJsonRpcBody('text/event-stream', ': ping - x\r\n\r\n')).toThrow(
      /no data frame/
    );
  });

  it('treats an absent content-type as JSON rather than guessing SSE', () => {
    expect(decodeJsonRpcBody(null, JSON.stringify(message))).toEqual(message);
  });

  /**
   * A CRLF blank line (`\r\n\r\n`) is ONE frame terminator, not a `\r` line
   * plus a real one. A reader that splits only on `\n` leaves that `\r`
   * attached as its own non-empty "line", which never equals `''`, so the
   * boundary between two frames is missed and their data gets joined into one
   * blob instead of returning the first frame alone. Two back-to-back frames
   * with genuinely different JSON make that observable: merging them yields
   * two top-level JSON values back to back, which `JSON.parse` rejects
   * outright (`Unexpected non-whitespace character after JSON`), where a
   * correct reader returns the first frame's message and never even looks at
   * the second.
   */
  it('treats a CRLF blank line as a single boundary, not two, between two frames', () => {
    const other = { jsonrpc: '2.0', id: 2, result: { structuredContent: { markdown: 'other' } } };
    const body =
      `event: message\r\ndata: ${JSON.stringify(message)}\r\n\r\n` +
      `event: message\r\ndata: ${JSON.stringify(other)}\r\n\r\n`;
    expect(decodeJsonRpcBody('text/event-stream', body)).toEqual(message);
  });

  /**
   * Under SSE the body IS the conversion: a frame that accumulated `data:`
   * lines but never reached its terminating blank line means the stream cut
   * off before the server finished, not that the server had nothing to say.
   * Returning that partial accumulation as success would be a clean,
   * well-formed, wrong answer — the same failure class as the empty 200 of
   * 2026-08-07 — so this must throw, distinguishably from "no data frame".
   */
  it('throws on a stream that ends mid-frame, before its terminating blank line', () => {
    const body = `event: message\r\ndata: ${JSON.stringify(message)}`;
    expect(() => decodeJsonRpcBody('text/event-stream', body)).toThrow(/truncat/);
  });
});
