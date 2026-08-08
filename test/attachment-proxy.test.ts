import { describe, it, expect, vi } from 'vitest';
import { AttachmentProxyClient } from '../src/lib/attachment-proxy.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn(), verbose: vi.fn() },
  enableConsoleLogging: vi.fn(),
}));

/**
 * `fetch` is the only thing mocked in this file, and that is the point.
 *
 * On 2026-08-07 a defect shipped because every route test mocked the layer
 * below the wire and gave it a truth the real client could not tell. Mocking
 * `AttachmentProxyClient` from the tool layer would repeat that exactly: the
 * framing, the headers and the error mapping would never run. So the fake is a
 * `fetch` returning real `Response` objects, and the client under test is the
 * real one.
 */

/** A success payload shaped like the real tool's return value. */
function convertPayload(markdown: string) {
  return {
    markdown,
    next_offset: null,
    returned_chars: markdown.length,
    total_chars: markdown.length,
    content_status: 'ok',
    format: 'pdf',
    pages_total: 1,
    pages_returned: 1,
  };
}

/** A `tools/call` result as mcp puts it on the wire: both channels, no isError. */
export function okMessage(markdown: string) {
  const payload = convertPayload(markdown);
  return {
    jsonrpc: '2.0',
    id: 1,
    result: {
      content: [{ type: 'text', text: JSON.stringify(payload) }],
      structuredContent: payload,
    },
  };
}

export function sseResponse(message: unknown, status = 200): Response {
  return new Response(`event: message\r\ndata: ${JSON.stringify(message)}\r\n\r\n`, {
    status,
    headers: { 'content-type': 'text/event-stream; charset=utf-8' },
  });
}

export function jsonResponse(message: unknown, status = 200): Response {
  return new Response(JSON.stringify(message), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

/** A `fetch` stand-in that records what it was asked for. */
export function recordingFetch(responses: Response[]) {
  // RequestInit/RequestInfo are DOM-lib type-only names with no runtime
  // existence; core `no-undef` predates TypeScript and doesn't understand
  // type positions, so it flags them as undefined globals.
  // eslint-disable-next-line no-undef
  const calls: Array<{ url: string; init: RequestInit }> = [];
  const queue = [...responses];
  // eslint-disable-next-line no-undef
  const impl = (async (input: RequestInfo | URL, init?: RequestInit) => {
    calls.push({ url: String(input), init: init ?? {} });
    const next = queue.shift();
    if (!next) throw new Error('the client made more requests than the test queued');
    return next;
  }) as typeof fetch;
  return { impl, calls };
}

describe('AttachmentProxyClient.convertToMarkdown', () => {
  it('returns markdown from an SSE-framed 200', async () => {
    const { impl } = recordingFetch([sseResponse(okMessage('# Invoice\n\nTotal: $12'))]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    expect(await client.convertToMarkdown({ uri: 'http://m365:3001/attachment?t=abc' })).toEqual({
      ok: true,
      markdown: '# Invoice\n\nTotal: $12',
    });
  });

  it('returns markdown from a plain application/json 200', async () => {
    const { impl } = recordingFetch([jsonResponse(okMessage('# Invoice'))]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    expect(await client.convertToMarkdown({ uri: 'http://m365:3001/attachment?t=abc' })).toEqual({
      ok: true,
      markdown: '# Invoice',
    });
  });

  it('posts one stateless tools/call with no session handshake', async () => {
    const { impl, calls } = recordingFetch([sseResponse(okMessage('x'))]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    await client.convertToMarkdown({ uri: 'u', pages: '1-3', offset: 40, maxChars: 20000 });

    expect(calls).toHaveLength(1);
    expect(calls[0].url).toBe('http://proxy:8080/mcp');
    expect(calls[0].init.method).toBe('POST');
    expect(JSON.parse(String(calls[0].init.body))).toEqual({
      jsonrpc: '2.0',
      id: 1,
      method: 'tools/call',
      params: {
        name: 'convert_to_markdown',
        arguments: { uri: 'u', pages: '1-3', offset: 40, max_chars: 20000 },
      },
    });
  });

  it('sends the Accept header Streamable HTTP requires', async () => {
    // mcp answers 406 unless the POST accepts *both* media types. A client that
    // sends only application/json never gets a document at all.
    const { impl, calls } = recordingFetch([sseResponse(okMessage('x'))]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    await client.convertToMarkdown({ uri: 'u' });

    const headers = calls[0].init.headers as Record<string, string>;
    expect(headers.Accept).toBe('application/json, text/event-stream');
    expect(headers['Content-Type']).toBe('application/json');
  });

  it('reports a result carrying no markdown as an error rather than empty success', async () => {
    // The empty-200 shape, one layer up: a well-formed result with nothing in
    // it must never be handed to the model as a successfully converted blank.
    const { impl } = recordingFetch([
      sseResponse({ jsonrpc: '2.0', id: 1, result: { content: [], structuredContent: {} } }),
    ]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    const result = await client.convertToMarkdown({ uri: 'u' });
    expect(result).toEqual({
      ok: false,
      code: 'proxy_error',
      message: expect.stringContaining('no markdown'),
    });
  });
});

/** A coded tool error as `_error_result` puts it on the wire: both channels, isError true. */
export function errorMessage(code: string, message: string, detail: Record<string, unknown> = {}) {
  const payload = { code, message, detail };
  return {
    jsonrpc: '2.0',
    id: 1,
    result: {
      content: [{ type: 'text', text: JSON.stringify(payload) }],
      structuredContent: payload,
      isError: true,
    },
  };
}

describe('AttachmentProxyClient error mapping', () => {
  it('passes a contract code through unchanged', async () => {
    const { impl } = recordingFetch([
      sseResponse(errorMessage('unsupported_format', 'This server cannot convert 7z archives.')),
    ]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    expect(await client.convertToMarkdown({ uri: 'u' })).toEqual({
      ok: false,
      code: 'unsupported_format',
      message: 'This server cannot convert 7z archives.',
    });
  });

  it('passes an unknown proxy code through as proxy_error with the raw code intact', async () => {
    // The proxy's own vocabulary is not ours to promise. `busy` is real and is
    // outside the spec's table; flattening it to conversion_failed would tell
    // the agent "report this" when the right answer is "try again".
    const { impl } = recordingFetch([
      sseResponse(errorMessage('busy', 'Every conversion worker is checked out.')),
    ]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    const result = await client.convertToMarkdown({ uri: 'u' });
    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('unreachable');
    expect(result.code).toBe('proxy_error');
    // Verbatim and first, so it survives truncation and greps cleanly.
    expect(result.message.startsWith('busy: ')).toBe(true);
    expect(result.message).toContain('Every conversion worker is checked out.');
  });

  it('carries the upstream status on a fetch_failed', async () => {
    const { impl } = recordingFetch([
      sseResponse(
        errorMessage('fetch_failed', 'The document could not be fetched.', { status: 404 })
      ),
    ]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    const result = await client.convertToMarkdown({ uri: 'u' });
    expect(result).toEqual({
      ok: false,
      code: 'fetch_failed',
      message: 'The document could not be fetched. (upstream status 404)',
    });
  });

  it('recovers the code from the text channel when structuredContent is absent', async () => {
    // A conforming proxy that is not this one may put the payload only in the
    // text block. The code still has to survive.
    const payload = { code: 'too_large', message: 'The document exceeds the size limit.' };
    const { impl } = recordingFetch([
      sseResponse({
        jsonrpc: '2.0',
        id: 1,
        result: { content: [{ type: 'text', text: JSON.stringify(payload) }], isError: true },
      }),
    ]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    expect(await client.convertToMarkdown({ uri: 'u' })).toEqual({
      ok: false,
      code: 'too_large',
      message: 'The document exceeds the size limit.',
    });
  });

  it('codes an error that arrives with no code at all', async () => {
    const { impl } = recordingFetch([
      sseResponse({
        jsonrpc: '2.0',
        id: 1,
        result: { content: [{ type: 'text', text: 'exploded' }], isError: true },
      }),
    ]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    const result = await client.convertToMarkdown({ uri: 'u' });
    expect(result).toEqual({
      ok: false,
      code: 'proxy_error',
      message: expect.stringContaining('no code'),
    });
  });

  it('maps a JSON-RPC envelope error to proxy_error rather than pretending it converted', async () => {
    // What a 401 or an unserved method looks like: an error member on the
    // envelope, no result at all.
    const { impl } = recordingFetch([
      jsonResponse(
        { jsonrpc: '2.0', id: null, error: { code: -32600, message: 'Unauthorized' } },
        401
      ),
    ]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    const result = await client.convertToMarkdown({ uri: 'u' });
    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('unreachable');
    expect(result.code).toBe('proxy_error');
    expect(result.message).toContain('-32600');
    expect(result.message).toContain('Unauthorized');
  });

  it('codes a body it cannot read at all instead of throwing at the caller', async () => {
    const { impl } = recordingFetch([
      new Response('<html>502 Bad Gateway</html>', {
        status: 502,
        headers: { 'content-type': 'text/html' },
      }),
    ]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    const result = await client.convertToMarkdown({ uri: 'u' });
    expect(result).toEqual({
      ok: false,
      code: 'proxy_error',
      message: expect.stringContaining('502'),
    });
  });
});

import { afterEach } from 'vitest';
import logger from '../src/logger.js';

describe('AttachmentProxyClient transport failures', () => {
  afterEach(() => {
    vi.mocked(logger.debug).mockClear();
  });

  it('codes a connection refusal as proxy_unreachable instead of throwing', async () => {
    // Exactly what undici raises when nothing is listening: a TypeError whose
    // cause carries the errno. The client must survive both layers.
    const refused = new TypeError('fetch failed');
    (refused as { cause?: unknown }).cause = Object.assign(
      new Error('connect ECONNREFUSED 192.168.128.9:8080'),
      { code: 'ECONNREFUSED' }
    );
    const impl = (async () => {
      throw refused;
    }) as typeof fetch;
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    const result = await client.convertToMarkdown({ uri: 'u' });
    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('unreachable');
    expect(result.code).toBe('proxy_unreachable');
    expect(result.message).toContain('ECONNREFUSED');
  });

  it('logs a proxy_unreachable at DEBUG level (not warn) with the url and the elapsed time', async () => {
    // Downgraded from warn: `read-document` in graph-tools.ts is the layer that
    // knows the attempt number and owns the operator-facing WARN now (worded
    // "attempt 1 of 2" / "still unreachable after retry"). A warn at both
    // layers would double every real failure into two differently-worded WARN
    // lines from two clocks -- see attachment-proxy-read-document.test.ts's
    // "logs exactly one WARN" assertion for the layer that must NOT also warn.
    const impl = (async () => {
      throw new TypeError('fetch failed');
    }) as typeof fetch;
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    await client.convertToMarkdown({ uri: 'u' });

    expect(logger.warn).not.toHaveBeenCalled();
    expect(logger.debug).toHaveBeenCalledTimes(1);
    const line = String(vi.mocked(logger.debug).mock.calls[0][0]);
    expect(line).toContain('http://proxy:8080/mcp');
    expect(line).toMatch(/\d+ ms/);
  });

  it('gives up at the timeout and codes it proxy_unreachable', async () => {
    vi.useFakeTimers();
    try {
      // A proxy that accepted the connection and then went quiet -- the wedged
      // worker case. The mock aborts when the client's signal fires, which is
      // the only thing that can end this promise.
      // eslint-disable-next-line no-undef
      const impl = ((_input: RequestInfo | URL, init?: RequestInit) =>
        new Promise<Response>((_resolve, reject) => {
          init?.signal?.addEventListener('abort', () => {
            const error = new Error('This operation was aborted');
            error.name = 'AbortError';
            reject(error);
          });
        })) as typeof fetch;
      const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

      const pending = client.convertToMarkdown({ uri: 'u' });
      await vi.advanceTimersByTimeAsync(60_000);
      const result = await pending;

      expect(result.ok).toBe(false);
      if (result.ok) throw new Error('unreachable');
      expect(result.code).toBe('proxy_unreachable');
      expect(result.message).toContain('60000 ms');
    } finally {
      vi.useRealTimers();
    }
  });

  it('honours an explicit timeoutMs over the 60 s default', async () => {
    vi.useFakeTimers();
    try {
      // eslint-disable-next-line no-undef
      const impl = ((_input: RequestInfo | URL, init?: RequestInit) =>
        new Promise<Response>((_resolve, reject) => {
          init?.signal?.addEventListener('abort', () => {
            const error = new Error('This operation was aborted');
            error.name = 'AbortError';
            reject(error);
          });
        })) as typeof fetch;
      const client = new AttachmentProxyClient({
        url: 'http://proxy:8080/mcp',
        timeoutMs: 1500,
        fetchImpl: impl,
      });

      const pending = client.convertToMarkdown({ uri: 'u' });
      await vi.advanceTimersByTimeAsync(1500);

      expect((await pending).ok).toBe(false);
    } finally {
      vi.useRealTimers();
    }
  });

  it('does not leave the timer armed after a fast success', async () => {
    vi.useFakeTimers();
    try {
      const { impl } = recordingFetch([sseResponse(okMessage('# quick'))]);
      const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

      expect(await client.convertToMarkdown({ uri: 'u' })).toEqual({
        ok: true,
        markdown: '# quick',
      });
      // A timer still pending here would keep a 60 s handle alive per call.
      expect(vi.getTimerCount()).toBe(0);
    } finally {
      vi.useRealTimers();
    }
  });

  it('times out on a stalled body even though the response headers arrived immediately', async () => {
    // The property the whole task exists for: under SSE the body IS the
    // conversion, so a client that only bounds the time-to-headers would
    // never observe this failure. The mock resolves the outer fetch() call
    // right away with a real, headers-bearing Response, then stalls its body
    // read (`.text()`) until the abort fires -- proving the timer's abort
    // signal reaches the body consumption step, not merely the connect step.
    vi.useFakeTimers();
    try {
      let bodyAborted = false;
      // eslint-disable-next-line no-undef
      const impl = ((_input: RequestInfo | URL, init?: RequestInit) => {
        const response = new Response('', {
          status: 200,
          headers: { 'content-type': 'text/event-stream' },
        });
        const stalledText = new Promise<string>((_resolve, reject) => {
          init?.signal?.addEventListener('abort', () => {
            bodyAborted = true;
            const error = new Error('This operation was aborted');
            error.name = 'AbortError';
            reject(error);
          });
        });
        response.text = () => stalledText;
        return Promise.resolve(response);
      }) as typeof fetch;
      const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

      const pending = client.convertToMarkdown({ uri: 'u' });
      // The headers resolve on the same microtask turn; nothing here waits on
      // real time, so this assertion would pass even with a headers-only
      // timeout. It's the assertions after the advance that distinguish them.
      await vi.advanceTimersByTimeAsync(60_000);
      const result = await pending;

      expect(bodyAborted).toBe(true);
      expect(result.ok).toBe(false);
      if (result.ok) throw new Error('unreachable');
      expect(result.code).toBe('proxy_unreachable');
      expect(result.message).toContain('60000 ms');
    } finally {
      vi.useRealTimers();
    }
  });
});

describe('AttachmentProxyClient authentication', () => {
  const saved = process.env.MS365_MCP_ATTACHMENT_PROXY_TOKEN;
  afterEach(() => {
    if (saved === undefined) delete process.env.MS365_MCP_ATTACHMENT_PROXY_TOKEN;
    else process.env.MS365_MCP_ATTACHMENT_PROXY_TOKEN = saved;
  });

  it('sends a bearer token when one is configured', async () => {
    process.env.MS365_MCP_ATTACHMENT_PROXY_TOKEN = 's3cret';
    const { impl, calls } = recordingFetch([sseResponse(okMessage('x'))]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    await client.convertToMarkdown({ uri: 'u' });

    expect((calls[0].init.headers as Record<string, string>).Authorization).toBe('Bearer s3cret');
  });

  it('sends no Authorization header at all when none is configured', async () => {
    delete process.env.MS365_MCP_ATTACHMENT_PROXY_TOKEN;
    const { impl, calls } = recordingFetch([sseResponse(okMessage('x'))]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    await client.convertToMarkdown({ uri: 'u' });

    // Absent, not empty: `Bearer ` is a credential a server may log as one.
    expect('Authorization' in (calls[0].init.headers as Record<string, string>)).toBe(false);
  });

  it('reads the token per call, so a rotated env var is picked up without a restart', async () => {
    process.env.MS365_MCP_ATTACHMENT_PROXY_TOKEN = 'first';
    const { impl, calls } = recordingFetch([
      sseResponse(okMessage('x')),
      sseResponse(okMessage('y')),
    ]);
    const client = new AttachmentProxyClient({ url: 'http://proxy:8080/mcp', fetchImpl: impl });

    await client.convertToMarkdown({ uri: 'u' });
    process.env.MS365_MCP_ATTACHMENT_PROXY_TOKEN = 'second';
    await client.convertToMarkdown({ uri: 'u' });

    expect((calls[0].init.headers as Record<string, string>).Authorization).toBe('Bearer first');
    expect((calls[1].init.headers as Record<string, string>).Authorization).toBe('Bearer second');
  });
});
