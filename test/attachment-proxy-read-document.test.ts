/**
 * read-document: from a Graph byte path to markdown, without a byte reaching
 * the model.
 *
 * The proxy client is a real object here only where the test is about the wire
 * (Task 15's proxy_unreachable cases dial a genuinely closed port). Where the
 * test is about the TOOL's branching, the client is stubbed at its own
 * interface -- which is not the trap the spec names. That trap was "nothing
 * exercises the wire at all"; the wire is covered by the proxy client's own
 * fetch-level suite and by the real-HTTP end-to-end test.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { registerGraphTools } from '../src/graph-tools.js';
import { AttachmentTicketStore } from '../src/lib/attachment-tickets.js';
import {
  configureAttachmentMinting,
  resetAttachmentMinting,
} from '../src/lib/attachment-minting.js';
import {
  configureAttachmentProxy,
  resetAttachmentProxy,
} from '../src/lib/attachment-proxy-runtime.js';
import { AttachmentProxyClient } from '../src/lib/attachment-proxy.js';
import type { ConvertRequest, ConvertResult } from '../src/lib/attachment-proxy.js';
import type GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn(), verbose: vi.fn() },
  enableConsoleLogging: vi.fn(),
}));

const MAIL_ATTACHMENT = '/me/messages/AAA/attachments/BBB/$value';

const URL_CONFIG = {
  base: 'http://m365-max-mcp:3001',
  key: 'shared-hmac-key',
  keyId: 'k1',
  ttlSeconds: 120,
};

/** A client stub typed as the real one, recording every request it is handed. */
function stubProxy(reply: ConvertResult): {
  client: AttachmentProxyClient;
  requests: ConvertRequest[];
} {
  const requests: ConvertRequest[] = [];
  const client = {
    convertToMarkdown: async (req: ConvertRequest): Promise<ConvertResult> => {
      requests.push(req);
      return reply;
    },
  } as unknown as AttachmentProxyClient;
  return { client, requests };
}

function fakeGraphClient(): GraphClient {
  return {
    makeRequest: vi.fn(async () => ({
      name: 'report.pdf',
      contentType: 'application/pdf',
      size: 195663,
    })),
  } as unknown as GraphClient;
}

/** The `uri` argument a `fetch` mock was actually handed, recovered from the JSON-RPC request body. */
function uriFromRequestBody(init: unknown): string {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const body = JSON.parse(String((init as any)?.body ?? '{}'));
  return body?.params?.arguments?.uri as string;
}

function jsonRpcResponse(message: unknown): Response {
  return new Response(JSON.stringify(message), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
}

/**
 * A REAL `AttachmentProxyClient`, talking to a `fetch` mock rather than to a
 * stub of the client's own interface. That distinction is the point: a stub
 * at the client's interface (as `stubProxy` above provides, correctly, for
 * every test about the tool's own branching) hands back exactly the string
 * the test author wrote and can never exercise `mapProxyError` /
 * `interpretJsonRpcMessage` in `attachment-proxy.ts` -- which is where an
 * untrusted proxy's text actually becomes this tool's `message`/`markdown`.
 * Mocking only `fetch` forces every byte the assertions below see through
 * that real derivation, the same way `attachment-proxy.test.ts` does for the
 * client's own suite.
 */
function realProxyClient(buildMessage: (requestedUri: string) => unknown): {
  client: AttachmentProxyClient;
  requestedUris: string[];
} {
  const requestedUris: string[] = [];
  const fetchImpl = (async (_input: unknown, init?: unknown) => {
    const uri = uriFromRequestBody(init);
    requestedUris.push(uri);
    return jsonRpcResponse(buildMessage(uri));
  }) as unknown as typeof fetch;
  const client = new AttachmentProxyClient({ url: 'http://docglean:8080/mcp', fetchImpl });
  return { client, requestedUris };
}

describe('read-document', () => {
  let store: AttachmentTicketStore;

  async function connect(): Promise<Client> {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    registerGraphTools(
      server,
      fakeGraphClient(),
      false,
      '^read-document$',
      false,
      undefined,
      false,
      [],
      undefined,
      true,
      true
    );
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await server.connect(serverTransport);
    const client = new Client({ name: 'test-client', version: '1.0.0' });
    await client.connect(clientTransport);
    return client;
  }

  async function call(
    client: Client,
    args: Record<string, unknown>
  ): Promise<{ isError: boolean; text: string }> {
    const result = (await client.callTool({ name: 'read-document', arguments: args })) as {
      isError?: boolean;
      content: Array<{ text: string }>;
    };
    return { isError: Boolean(result.isError), text: result.content[0].text };
  }

  beforeEach(() => {
    vi.clearAllMocks();
    store = new AttachmentTicketStore(120);
    configureAttachmentMinting({ store, config: URL_CONFIG });
  });

  afterEach(() => {
    resetAttachmentMinting();
    resetAttachmentProxy();
    vi.restoreAllMocks();
  });

  it('mints a ticket for a valid target and returns the markdown verbatim', async () => {
    const { client: proxy, requests } = stubProxy({ ok: true, markdown: '# Q3 report\n\nBody.' });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });

    const result = await call(await connect(), {
      target: MAIL_ATTACHMENT,
      pages: '1-3',
      maxChars: 8000,
    });

    expect(result.isError).toBe(false);
    expect(result.text).toBe('# Q3 report\n\nBody.');
    expect(requests).toHaveLength(1);
    expect(requests[0].pages).toBe('1-3');
    expect(requests[0].maxChars).toBe(8000);
    // A real, signed, redeemable URL built from the real store -- not a
    // placeholder, and not the raw Graph path.
    const uri = new URL(requests[0].uri);
    expect(uri.origin).toBe('http://m365-max-mcp:3001');
    expect(uri.pathname).toBe('/attachment');
    expect(uri.searchParams.get('t')).toMatch(/^[A-Za-z0-9_-]{20,}$/);
    expect(store.redemptionsLeft(uri.searchParams.get('t')!)).toBe(3);
  });

  it('refuses a target outside the mintable grammar as invalid_target', async () => {
    const { client: proxy, requests } = stubProxy({ ok: true, markdown: 'never reached' });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });

    const result = await call(await connect(), { target: '/me/messages/AAA' });

    expect(result.isError).toBe(true);
    const body = JSON.parse(result.text);
    expect(body.error).toBe('invalid_target');
    // Every failure carries the envelope, even the ones that know nothing.
    expect(body).toHaveProperty('name');
    expect(body).toHaveProperty('contentType');
    expect(body).toHaveProperty('size');
    // Nothing was minted and nothing was dialled.
    expect(store.size()).toBe(0);
    expect(requests).toHaveLength(0);
  });

  it('refuses an absolute URL as invalid_target', async () => {
    const { client: proxy } = stubProxy({ ok: true, markdown: 'never reached' });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });
    const result = await call(await connect(), {
      target: 'https://graph.microsoft.com/v1.0/me/messages/AAA/attachments/BBB/$value',
    });
    expect(JSON.parse(result.text).error).toBe('invalid_target');
  });

  it('mints a fresh ticket on every call rather than reusing one', async () => {
    const { client: proxy, requests } = stubProxy({ ok: true, markdown: 'ok' });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });
    const client = await connect();

    await call(client, { target: MAIL_ATTACHMENT });
    await call(client, { target: MAIL_ATTACHMENT });

    expect(requests).toHaveLength(2);
    const [first, second] = requests.map((r) => new URL(r.uri).searchParams.get('t'));
    expect(first).toBeTruthy();
    expect(second).toBeTruthy();
    // Two calls against the identical target mint two distinct capabilities --
    // the whole reason a fresh ticket is minted per call rather than a ticket
    // being cached and handed out again. Reusing one would make the 3-redemption
    // budget and 120s TTL leak across unrelated calls instead of bounding each.
    expect(first).not.toBe(second);
    expect(store.size()).toBe(2);
    expect(store.redemptionsLeft(first!)).toBe(3);
    expect(store.redemptionsLeft(second!)).toBe(3);
  });

  it('answers no_capacity when the ticket store is at its 256 cap', async () => {
    const { client: proxy, requests } = stubProxy({ ok: true, markdown: 'never reached' });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });
    for (let i = 0; i < 256; i++) store.mint(`/me/messages/M${i}/attachments/A/$value`, undefined);

    const result = await call(await connect(), { target: MAIL_ATTACHMENT });

    expect(result.isError).toBe(true);
    const body = JSON.parse(result.text);
    expect(body.error).toBe('no_capacity');
    expect(body).toHaveProperty('name');
    expect(requests).toHaveLength(0);
  });

  it('redacts the live ticket url out of a real proxy error message', async () => {
    // The real client, talking to a fetch mock whose error text echoes back
    // the exact uri it was asked to convert -- an entirely ordinary shape
    // ("could not fetch <uri>: 404"), not a hostile one. This is the case a
    // stub-based test cannot prove: it exercises the real mapProxyError /
    // interpretJsonRpcMessage derivation, not a literal a test author wrote.
    const { client: proxy, requestedUris } = realProxyClient((uri) => {
      const payload = {
        code: 'fetch_failed',
        message: `could not fetch ${uri}: upstream returned 404`,
        detail: { status: 404 },
      };
      return {
        jsonrpc: '2.0',
        id: 1,
        result: {
          content: [{ type: 'text', text: JSON.stringify(payload) }],
          structuredContent: payload,
          isError: true,
        },
      };
    });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });

    const result = await call(await connect(), { target: MAIL_ATTACHMENT });

    expect(requestedUris).toHaveLength(1);
    const ticketUrl = requestedUris[0];
    const ticketId = new URL(ticketUrl).searchParams.get('t')!;
    expect(ticketId).toMatch(/^[A-Za-z0-9_-]{20,}$/);

    expect(result.isError).toBe(true);
    const body = JSON.parse(result.text);
    expect(body.error).toBe('fetch_failed');
    // The whole point: neither the full signed URL nor the bare ticket id --
    // the actual credential the redemption route checks -- survives into the
    // tool's output, even though the proxy's own text carried both.
    expect(body.message).not.toContain(ticketUrl);
    expect(body.message).not.toContain(ticketId);
    expect(result.text).not.toContain(ticketUrl);
    expect(result.text).not.toContain(ticketId);
    expect(body.message).toContain('<attachment url redacted>');
  });

  it('redacts a bare ticket id even when the surrounding url is mangled or truncated', async () => {
    // A proxy that logs a parsed field rather than the raw request line would
    // echo the id without ever reproducing the exact signed URL string --
    // reordered query, dropped dgk/dgx/dgs, truncated at a delimiter. The id
    // itself is the credential the redemption route actually checks
    // (attachment-route.ts ignores dgk/dgx/dgs), so this has to be caught on
    // its own, not only as part of an exact URL match.
    const { client: proxy, requestedUris } = realProxyClient((uri) => {
      const ticketId = new URL(uri).searchParams.get('t');
      const payload = {
        code: 'proxy_error',
        message: `upstream 404 fetching /attachment (t=${ticketId}, dg* dropped by an intermediate proxy)`,
        detail: {},
      };
      return {
        jsonrpc: '2.0',
        id: 1,
        result: {
          content: [{ type: 'text', text: JSON.stringify(payload) }],
          structuredContent: payload,
          isError: true,
        },
      };
    });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });

    const result = await call(await connect(), { target: MAIL_ATTACHMENT });

    const ticketId = new URL(requestedUris[0]).searchParams.get('t')!;
    expect(result.text).not.toContain(ticketId);
    expect(result.text).toContain('<attachment url redacted>');
  });

  it('redacts an echoed ticket url out of a successful conversion, defensively', async () => {
    // Nothing stops a proxy from putting the request it received into the
    // markdown it returns, deliberately or by a debug echo left enabled --
    // and the ticket is exactly as live a credential there as in an error.
    const { client: proxy, requestedUris } = realProxyClient((uri) => {
      const payload = { markdown: `# doc\n\nconverted from: ${uri}` };
      return {
        jsonrpc: '2.0',
        id: 1,
        result: {
          content: [{ type: 'text', text: JSON.stringify(payload) }],
          structuredContent: payload,
        },
      };
    });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });

    const result = await call(await connect(), { target: MAIL_ATTACHMENT });

    expect(result.isError).toBe(false);
    const ticketUrl = requestedUris[0];
    const ticketId = new URL(ticketUrl).searchParams.get('t')!;
    expect(result.text).not.toContain(ticketUrl);
    expect(result.text).not.toContain(ticketId);
    expect(result.text).toContain('<attachment url redacted>');
    expect(result.text).toContain('# doc');
  });
});
