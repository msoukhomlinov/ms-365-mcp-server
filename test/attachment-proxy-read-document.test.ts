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
import { createServer, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
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
import logger from '../src/logger.js';

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

/**
 * The confirmed live defect and its fix: docglean does not sniff format, so
 * `read-document` converting correctly for every attachment kind depends on
 * `attachment-route.ts` serving a usable Content-Type -- and the only lever
 * this server has over that, for a mail/event attachment, is what it learned
 * about the target BEFORE minting. These tests pin what gets learned and
 * carried, not the route's precedence itself (that is
 * `attachment-route-content-type.test.ts`).
 */
describe('read-document mints with the content-type it already knows', () => {
  let store: AttachmentTicketStore;

  async function connect(graphClient: GraphClient): Promise<Client> {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    registerGraphTools(
      server,
      graphClient,
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

  function graphClientReturning(meta: Record<string, unknown> | null): GraphClient {
    return { makeRequest: async () => meta } as unknown as GraphClient;
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

  it('passes a specific probed content-type from Graph metadata into the mint', async () => {
    const { client: proxy } = stubProxy({ ok: true, markdown: 'ok' });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });
    const graphClient = graphClientReturning({
      name: 'report.pdf',
      contentType: 'application/pdf',
      size: 195663,
    });
    const mintSpy = vi.spyOn(store, 'mint');

    await call(await connect(graphClient), { target: MAIL_ATTACHMENT });

    expect(mintSpy).toHaveBeenCalledTimes(1);
    const probe = mintSpy.mock.calls[0]?.[3];
    expect(probe?.contentType).toBe('application/pdf');
    expect(probe?.name).toBe('report.pdf');
  });

  it('carries a directly-populated message/rfc822 content-type, verified against a real itemAttachment', async () => {
    // Verified live against the deployed mailbox: Graph's OWN metadata for the
    // itemAttachment "Sartre and de Beauvoir, Six Lectures at the RH" (a
    // nested forwarded message, 23,317 bytes) populates contentType directly
    // as "message/rfc822" -- read verbatim, no @odata.type inference involved.
    // This is the verified mechanism the fix actually depends on.
    const { client: proxy } = stubProxy({ ok: true, markdown: 'ok' });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });
    const graphClient = graphClientReturning({
      name: 'Sartre and de Beauvoir, Six Lectures at the RH',
      contentType: 'message/rfc822',
      size: 23317,
    });
    const mintSpy = vi.spyOn(store, 'mint');

    await call(await connect(graphClient), { target: MAIL_ATTACHMENT });

    expect(mintSpy).toHaveBeenCalledTimes(1);
    const probe = mintSpy.mock.calls[0]?.[3];
    expect(probe?.contentType).toBe('message/rfc822');
    expect(probe?.name).toBe('Sartre and de Beauvoir, Six Lectures at the RH');
  });

  it('does not invent a content-type when Graph leaves it null, even with an itemAttachment @odata.type present', async () => {
    // No @odata.type-gated default: no registered tool in this server's
    // endpoints.json reaches the single-entity attachment GET this probe
    // calls (only DELETE is registered for that path), so whether Graph
    // volunteers @odata.type there has never been verified live. A null
    // metadata contentType -- also verified live, on a DIFFERENT
    // itemAttachment ("Katusha") than the one above -- stays null here; the
    // route's own stream-Content-Type fallback is what still applies to that
    // case, unchanged from before this fix.
    const { client: proxy } = stubProxy({ ok: true, markdown: 'ok' });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });
    const graphClient = graphClientReturning({
      name: 'Katusha',
      contentType: null,
      size: 206268,
      '@odata.type': '#microsoft.graph.itemAttachment',
    });
    const mintSpy = vi.spyOn(store, 'mint');

    await call(await connect(graphClient), { target: MAIL_ATTACHMENT });

    const probe = mintSpy.mock.calls[0]?.[3];
    expect(probe?.contentType).toBeNull();
    expect(probe?.name).toBe('Katusha');
  });

  it('refuses a reference attachment before minting anything, with a clear error', async () => {
    // A referenceAttachment carries no bytes at all -- it is a link. Fetching
    // its /$value would either error confusingly or return something that is
    // not the linked file, so this must be refused up front rather than
    // spending a ticket and a proxy round trip on a conversion that cannot
    // succeed.
    //
    // UNVERIFIED LIVE: no referenceAttachment specimen was ever found in the
    // target mailbox despite a broad search (see the content-type report),
    // and this detection depends on the same @odata.type annotation the
    // itemAttachment default above turned out NOT to be able to rely on --
    // no registered tool reaches the single-entity GET this probe calls, so
    // whether Graph actually sends @odata.type here has never been observed
    // live. This mocked test proves the WIRING (refuse-before-mint, given the
    // annotation) is correct; it does not prove the annotation arrives in
    // production. If it never does, this refusal simply never fires and
    // behaviour for that (unconfirmed) case is unchanged from before this fix
    // -- see probeMailEventAttachment's docstring in graph-tools.ts.
    const { client: proxy, requests } = stubProxy({ ok: true, markdown: 'never reached' });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });
    const graphClient = graphClientReturning({
      name: 'Shared design doc',
      contentType: null,
      size: 48213,
      '@odata.type': '#microsoft.graph.referenceAttachment',
      sourceUrl: 'https://contoso.sharepoint.com/:w:/link',
    });
    const mintSpy = vi.spyOn(store, 'mint');

    const result = await call(await connect(graphClient), { target: MAIL_ATTACHMENT });

    expect(result.isError).toBe(true);
    const body = JSON.parse(result.text);
    // Actionable, not a bare "conversion failed": names what this actually is.
    expect(body.error).toBe('reference_attachment');
    expect(body.message).toMatch(/reference|link/i);
    expect(body.name).toBe('Shared design doc');
    expect(mintSpy).not.toHaveBeenCalled();
    expect(requests).toHaveLength(0);
    expect(store.size()).toBe(0);
  });

  it('keeps minting usable when the probe itself fails, defaulting to no information', async () => {
    const { client: proxy, requests } = stubProxy({ ok: true, markdown: 'ok despite no probe' });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });
    const graphClient = {
      makeRequest: async () => {
        throw new Error('Graph 404');
      },
    } as unknown as GraphClient;
    const mintSpy = vi.spyOn(store, 'mint');

    const result = await call(await connect(graphClient), { target: MAIL_ATTACHMENT });

    expect(result.isError).toBe(false);
    expect(requests).toHaveLength(1);
    const probe = mintSpy.mock.calls[0]?.[3];
    expect(probe?.contentType).toBeNull();
    expect(probe?.name).toBeNull();
  });

  it('does not probe at all for a target that is not a mail/event attachment', async () => {
    const { client: proxy } = stubProxy({ ok: true, markdown: 'ok' });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });
    const makeRequest = vi.fn(async () => ({ name: 'x', contentType: 'x', size: 1 }));
    const graphClient = { makeRequest } as unknown as GraphClient;

    await call(await connect(graphClient), { target: '/drives/DRIVE1/items/ITEM1/content' });

    expect(makeRequest).not.toHaveBeenCalled();
  });
});

/** A port nothing is listening on: bound to learn the number, then released. */
async function reserveClosedPort(): Promise<number> {
  const holder = await new Promise<Server>((resolve) => {
    const s = createServer();
    s.listen(0, '127.0.0.1', () => resolve(s));
  });
  const port = (holder.address() as AddressInfo).port;
  await new Promise<void>((resolve) => holder.close(() => resolve()));
  return port;
}

describe('read-document failure handling', () => {
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

  async function call(client: Client, args: Record<string, unknown>): Promise<string> {
    const result = (await client.callTool({ name: 'read-document', arguments: args })) as {
      content: Array<{ text: string }>;
    };
    return result.content[0].text;
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

  it('carries the real name, contentType and size on a refusal', async () => {
    const { client: proxy } = stubProxy({
      ok: false,
      code: 'too_large',
      message: 'document exceeds the configured size limit',
    });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });

    const body = JSON.parse(await call(await connect(), { target: MAIL_ATTACHMENT }));

    expect(body.error).toBe('too_large');
    // The whole point of the envelope: the agent can still tell the user what
    // it could not read.
    expect(body.name).toBe('report.pdf');
    expect(body.contentType).toBe('application/pdf');
    expect(body.size).toBe(195663);
  });

  it('passes a code outside the contract through as proxy_error, verbatim', async () => {
    const { client: proxy } = stubProxy({
      ok: false,
      code: 'ocr_backend_unavailable',
      message: 'the OCR worker pool is not accepting work',
    });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });

    const body = JSON.parse(await call(await connect(), { target: MAIL_ATTACHMENT }));

    // Flattening this into conversion_failed would promise a vocabulary that is
    // not ours -- the proxy contract names a code, not a code list.
    expect(body.error).toBe('proxy_error');
    expect(body.proxyCode).toBe('ocr_backend_unavailable');
    expect(body.message).toBe('the OCR worker pool is not accepting work');
  });

  it('reports an unreachable proxy AND logs it at warn with the URL and elapsed ms -- exactly once per attempt, never twice', async () => {
    const deadPort = await reserveClosedPort();
    const url = `http://127.0.0.1:${deadPort}/mcp`;
    // The real client against a real closed socket. Nothing is stubbed: the one
    // code whose entire meaning is "the network failed" is not worth asserting
    // against a stub that decided to say so.
    const proxyClient = new AttachmentProxyClient({ url, timeoutMs: 2000 });
    // A dead port stays dead for the retry too, so this call legitimately makes
    // TWO attempts (see 'retries an unreachable proxy once' below) and this spy
    // is how the assertions distinguish "one warn per attempt" -- correct --
    // from "one warn per attempt per layer" -- the regression this test exists
    // to catch -- without hard-coding a count that would be wrong the moment a
    // retry is involved.
    const convertSpy = vi.spyOn(proxyClient, 'convertToMarkdown');
    configureAttachmentProxy({ client: proxyClient, url });

    const body = JSON.parse(await call(await connect(), { target: MAIL_ATTACHMENT }));

    expect(body.error).toBe('proxy_unreachable');
    expect(body.name).toBe('report.pdf');

    // Case-insensitive on purpose. `AttachmentProxyClient` used to log the
    // identical condition itself, spelled "[ATTACHMENT PROXY]" in capitals; a
    // case-sensitive `l.includes('proxy')` (the previous version of this
    // assertion) never saw that line and so never caught the duplication. The
    // client now logs at debug, not warn (see attachment-proxy.ts and its own
    // "logs a proxy_unreachable at DEBUG level (not warn)" test), so nothing
    // it emits should appear here at all.
    const allWarns = (logger.warn as ReturnType<typeof vi.fn>).mock.calls.map((c) => String(c[0]));
    const matching = allWarns.filter((l) => /proxy/i.test(l));
    // One WARN per attempt actually made -- not one per attempt per layer.
    // A future re-introduction of the client's own warn would make this
    // fail: convertSpy.mock.calls.length stays 2 (the retry is unaffected),
    // but matching.length would jump to 4.
    expect(matching).toHaveLength(convertSpy.mock.calls.length);
    // No two of those lines are identical text, which a duplicate emitter
    // (the same condition logged twice by two layers with the same wording)
    // would produce.
    expect(new Set(matching).size).toBe(matching.length);
    for (const line of matching) {
      expect(line).toContain(url);
      expect(line, 'the warn line must state elapsed time').toMatch(/\d+ms/);
    }
  });

  it('retries an unreachable proxy once, with a FRESH ticket', async () => {
    const deadPort = await reserveClosedPort();
    const url = `http://127.0.0.1:${deadPort}/mcp`;
    configureAttachmentProxy({
      client: new AttachmentProxyClient({ url, timeoutMs: 2000 }),
      url,
    });
    const mintSpy = vi.spyOn(store, 'mint');

    await call(await connect(), { target: MAIL_ATTACHMENT });

    // Two attempts, two mints, two DIFFERENT ids. Reusing the ticket is the
    // cheaper-looking mistake: a failed fetch still spends a redemption, and the
    // proxy may have spent one or more before failing, so the retry could meet a
    // 404 that has nothing to do with why the first attempt failed.
    expect(mintSpy).toHaveBeenCalledTimes(2);
    const ids = mintSpy.mock.results.map((r) => (r.value as { id: string }).id);
    expect(ids[0]).not.toBe(ids[1]);
  });

  it('does not retry a refusal the proxy actually answered', async () => {
    let calls = 0;
    const client = {
      convertToMarkdown: async () => {
        calls++;
        return { ok: false as const, code: 'password_required', message: 'encrypted PDF' };
      },
    } as unknown as AttachmentProxyClient;
    configureAttachmentProxy({ client, url: 'http://docglean:8080/mcp' });

    const body = JSON.parse(await call(await connect(), { target: MAIL_ATTACHMENT }));

    expect(body.error).toBe('password_required');
    // A proxy that answered is not a transient network failure. Retrying would
    // double every conversion cost for an answer that will not change.
    expect(calls).toBe(1);
  });

  it('redacts the SECOND ticket id when the retry succeeds and echoes it', async () => {
    // The retry mints a fresh ticket distinct from the first (proved above).
    // Redaction has to follow it: a naive implementation that redacts only
    // once, at the end, against whatever ticket was minted last, happens to
    // get this case right by coincidence -- but one that captured the FIRST
    // ticket's id before the retry and redacted only that would leak the
    // second ticket's id straight into the model's context on exactly the
    // path this test exercises.
    const requests: string[] = [];
    let calls = 0;
    const client = {
      convertToMarkdown: async (req: ConvertRequest) => {
        calls++;
        requests.push(req.uri);
        if (calls === 1) {
          return { ok: false as const, code: 'proxy_unreachable', message: 'timeout' };
        }
        return { ok: true as const, markdown: `converted from: ${req.uri}` };
      },
    } as unknown as AttachmentProxyClient;
    configureAttachmentProxy({ client, url: 'http://docglean:8080/mcp' });

    const text = await call(await connect(), { target: MAIL_ATTACHMENT });

    expect(calls).toBe(2);
    expect(requests[0]).not.toBe(requests[1]);
    const secondTicketUrl = requests[1];
    const secondTicketId = new URL(secondTicketUrl).searchParams.get('t')!;
    expect(text).not.toContain(secondTicketUrl);
    expect(text).not.toContain(secondTicketId);
    expect(text).toContain('<attachment url redacted>');
    expect(text).toContain('converted from:');
  });

  it('never lets a failed metadata probe replace the error the caller hit', async () => {
    const { client: proxy } = stubProxy({
      ok: false,
      code: 'conversion_failed',
      message: 'converter returned no text',
    });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });

    const server = new McpServer({ name: 'test', version: '1.0.0' });
    registerGraphTools(
      server,
      {
        makeRequest: vi.fn(async () => {
          throw new Error('Graph 404');
        }),
      } as unknown as GraphClient,
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

    const body = JSON.parse(await call(client, { target: MAIL_ATTACHMENT }));

    expect(body.error).toBe('conversion_failed');
    expect(body.name).toBeNull();
    expect(body.contentType).toBeNull();
    expect(body.size).toBeNull();
  });
});
