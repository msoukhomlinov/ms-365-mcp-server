/**
 * The whole path, with only Graph itself faked.
 *
 * The defect that shipped on 2026-08-07 is the reason this exists in this shape:
 * every route test mocked `downloadStream` and handed it a truthful
 * `contentLength`, so the mock told a truth the real client could not and the
 * derivation was never exercised. The same trap is live here -- if the proxy leg
 * is a stub, the proxy leg is never tested.
 *
 * So: a real HTTP proxy that really dials the URL it is given, a real ticket
 * minted by the running server's own store, and a real redemption through the
 * real :3001 listener. The only stub is Graph's byte stream, which is the one
 * thing this test cannot own -- and even that stub's PRIMARY fixture is
 * `contentLength: null`, matching what Graph actually sends for a `/$value`
 * fetch, so the null-length branch through `parseContentLengthHeader` and the
 * attachment route's header guard is the one this test exercises by default,
 * not a truthful-length shortcut around it. A secondary test covers the
 * truthful-length case too, since Graph is not forbidden from ever sending one.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createServer, type IncomingMessage, type Server, type ServerResponse } from 'node:http';
import type { AddressInfo } from 'node:net';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import MicrosoftGraphServer from '../src/server.js';
import GraphClient from '../src/graph-client.js';
import { registerGraphTools } from '../src/graph-tools.js';
import type AuthManager from '../src/auth.js';
import type { CommandOptions } from '../src/cli.js';
import { resetAttachmentMinting } from '../src/lib/attachment-minting.js';
import { resetAttachmentProxy } from '../src/lib/attachment-proxy-runtime.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn(), verbose: vi.fn() },
  enableConsoleLogging: vi.fn(),
}));

const MAIL_ATTACHMENT = '/me/messages/AAA/attachments/BBB/$value';
const DOCUMENT_BYTES = 'THE REAL PDF BYTES';

async function reserveFreePorts(count: number, host = '127.0.0.1'): Promise<number[]> {
  const holders = await Promise.all(
    Array.from(
      { length: count },
      () =>
        new Promise<Server>((resolve) => {
          const s = createServer();
          s.listen(0, host, () => resolve(s));
        })
    )
  );
  const ports = holders.map((s) => (s.address() as AddressInfo).port);
  await Promise.all(holders.map((s) => new Promise<void>((resolve) => s.close(() => resolve()))));
  return ports;
}

function fakeAuthManager(): AuthManager {
  return {
    isOAuthModeEnabled: () => false,
    isMultiAccount: async () => false,
    listAccounts: async () => [],
    getToken: async () => 'SERVER_OWN_TOKEN',
    getTokenForAccount: async () => 'SERVER_OWN_TOKEN',
    setOAuthToken: async () => {},
  } as unknown as AuthManager;
}

/**
 * A minimal stateless MCP proxy: one POST, no session handshake, and a
 * convert_to_markdown that ACTUALLY FETCHES the uri it was handed. Records every
 * uri, every pull-back status and every pull-back `content-length` header so
 * the test can assert the loop closed AND that the header the attachment route
 * declared was never a lie -- `content-length: 0` on a non-empty body being
 * the exact defect this whole feature exists to have caught.
 *
 * The response is shaped exactly like the real contract
 * (`AttachmentProxyClient.interpretJsonRpcMessage`, see test/attachment-proxy.test.ts's
 * `okMessage` helper): a JSON-RPC result whose text-channel content is a JSON
 * OBJECT carrying a `markdown` field, not raw markdown text. A fake that sent
 * raw text here would look like a legitimate proxy failure to the real client
 * and would silently coax this test into "fixing" the client instead of the fake.
 */
function startFakeProxy(
  port: number,
  seen: {
    uris: string[];
    statuses: number[];
    auth: (string | undefined)[];
    contentLengths: (string | null)[];
  }
): Promise<Server> {
  const handler = async (req: IncomingMessage, res: ServerResponse): Promise<void> => {
    const chunks: Buffer[] = [];
    for await (const chunk of req) chunks.push(chunk as Buffer);
    const body = JSON.parse(Buffer.concat(chunks).toString('utf8')) as {
      id?: number | string;
      params?: { name?: string; arguments?: Record<string, unknown> };
    };
    seen.auth.push(req.headers.authorization);
    const args = body.params?.arguments ?? {};
    const uri = String(args.uri ?? '');
    seen.uris.push(uri);

    const pulled = await fetch(uri);
    seen.statuses.push(pulled.status);
    seen.contentLengths.push(pulled.headers.get('content-length'));
    const text = await pulled.text();
    const markdown = `# report.pdf\n\n${text}\n`;
    const payload = { markdown };

    res.writeHead(200, { 'content-type': 'application/json' });
    res.end(
      JSON.stringify({
        jsonrpc: '2.0',
        id: body.id ?? 1,
        result: {
          content: [{ type: 'text', text: JSON.stringify(payload) }],
          structuredContent: payload,
          isError: false,
        },
      })
    );
  };
  return new Promise((resolve) => {
    const server = createServer((req, res) => {
      handler(req, res).catch(() => {
        if (!res.headersSent) res.writeHead(500);
        res.end();
      });
    });
    server.listen(port, '127.0.0.1', () => resolve(server));
  });
}

describe('--attachment-proxy end to end', () => {
  const savedEnv = { ...process.env };
  let started: MicrosoftGraphServer[] = [];
  let proxy: Server | null = null;
  let seen: {
    uris: string[];
    statuses: number[];
    auth: (string | undefined)[];
    contentLengths: (string | null)[];
  };
  let mcpPort: number;
  let attachmentPort: number;
  let proxyPort: number;

  /**
   * Re-stubs `downloadStream` with a given `contentLength`. Exposed as a
   * function, not inlined, so a test can call it a second time to swap the
   * fixture mid-test (see the "truthful length" case below) without touching
   * anything else `beforeEach` set up.
   *
   * `null` is the DEFAULT and the one `beforeEach` installs on its own,
   * because it is what real Graph actually sends for a `/$value` attachment
   * fetch (see the docstring on `parseContentLengthHeader`,
   * src/graph-client.ts:50-61: Graph does not send `content-length` on this
   * endpoint at all). A fixture that always hands back a truthful positive
   * number -- as this file originally did -- can never exercise the null
   * branch that `parseContentLengthHeader` and the route's header guard
   * (src/attachment-route.ts:120-126) exist for, which is exactly the
   * 2026-08-07 shape recurring inside the test meant to catch it.
   */
  function stubDownloadStream(contentLength: number | null): void {
    vi.spyOn(GraphClient.prototype, 'downloadStream').mockImplementation(async () => ({
      body: new ReadableStream({
        start(controller) {
          controller.enqueue(new TextEncoder().encode(DOCUMENT_BYTES));
          controller.close();
        },
      }) as never,
      contentType: 'application/pdf',
      contentLength,
      contentDisposition: 'attachment; filename="report.pdf"',
    }));
  }

  beforeEach(async () => {
    vi.clearAllMocks();
    seen = { uris: [], statuses: [], auth: [], contentLengths: [] };
    [mcpPort, attachmentPort, proxyPort] = await reserveFreePorts(3);

    process.env.MS365_MCP_RATE_LIMIT_DISABLED = 'true';
    process.env.MS365_MCP_ATTACHMENT_URL_KEY = 'shared-hmac-key';
    process.env.MS365_MCP_ATTACHMENT_URL_BASE = `http://127.0.0.1:${attachmentPort}`;
    process.env.MS365_MCP_ATTACHMENT_PROXY_TOKEN = 'proxy-bearer-token';

    stubDownloadStream(null);

    proxy = await startFakeProxy(proxyPort, seen);

    const options: CommandOptions = {
      http: `127.0.0.1:${mcpPort}`,
      trustProxyAuth: true,
      enableAttachmentUrls: true,
      attachmentPort: String(attachmentPort),
      attachmentProxy: `http://127.0.0.1:${proxyPort}/mcp`,
    };
    const server = new MicrosoftGraphServer(fakeAuthManager(), options);
    await server.initialize('0.0.0-test');
    started.push(server);
    await server.start();
  });

  afterEach(async () => {
    for (const server of started) await server.stop();
    started = [];
    await new Promise<void>((resolve) => (proxy ? proxy.close(() => resolve()) : resolve()));
    proxy = null;
    resetAttachmentMinting();
    resetAttachmentProxy();
    vi.restoreAllMocks();
    process.env = { ...savedEnv };
  });

  /**
   * An MCP client over the module singletons the running server configured. The
   * agent->m365 leg is in-memory on purpose: the three legs under test are the
   * mint, the proxy POST and the :3001 pull-back, and the running server's real
   * store and real client are what serve all three.
   */
  async function agent(): Promise<Client> {
    const mcp = new McpServer({ name: 'test', version: '1.0.0' });
    registerGraphTools(
      mcp,
      new GraphClient(fakeAuthManager(), {} as never),
      false,
      undefined,
      false,
      fakeAuthManager(),
      false,
      [],
      undefined,
      true,
      true
    );
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await mcp.connect(serverTransport);
    const client = new Client({ name: 'test-client', version: '1.0.0' });
    await client.connect(clientTransport);
    return client;
  }

  it('turns a Graph attachment path into markdown, through a real proxy and a real :3001 fetch, with the length Graph actually sends', async () => {
    // downloadStream is stubbed with contentLength: null by beforeEach -- the
    // normal production shape for a /$value fetch, where Graph never sends a
    // content-length header at all.
    const client = await agent();
    const result = (await client.callTool({
      name: 'read-document',
      arguments: { target: MAIL_ATTACHMENT },
    })) as { isError?: boolean; content: Array<{ text: string }> };

    expect(result.isError).toBeFalsy();
    expect(result.content[0].text).toBe(`# report.pdf\n\n${DOCUMENT_BYTES}\n`);

    // A non-empty body actually crossed the wire -- the exact failure class
    // that shipped on 2026-08-07 was a clean, well-formed, EMPTY 200.
    expect(result.content[0].text.length).toBeGreaterThan(DOCUMENT_BYTES.length);

    // The proxy was really dialled, once, with a real signed ticket URL...
    expect(seen.uris).toHaveLength(1);
    const uri = new URL(seen.uris[0]);
    expect(uri.port).toBe(String(attachmentPort));
    expect(uri.pathname).toBe('/attachment');
    expect(uri.searchParams.get('t')).toBeTruthy();
    expect(uri.searchParams.get('dgs')).toBeTruthy();
    // ...and its pull-back really reached the :3001 listener.
    expect(seen.statuses).toEqual([200]);
    // The bearer credential rode along.
    expect(seen.auth[0]).toBe('Bearer proxy-bearer-token');

    // The load-bearing assertion for this fixture: with no content-length
    // from Graph, the route must OMIT the header rather than declare a lying
    // one. It must never, under any circumstance, be the literal string '0'
    // on a body that is not empty -- that is the exact 2026-08-07 defect.
    expect(seen.contentLengths[0]).toBeNull();
    expect(seen.contentLengths[0]).not.toBe('0');

    // Neither the minted ticket id nor the URL it rode in on leaked into the
    // tool's own output -- Task 14's redactAttachmentSecrets, proved end to end.
    const ticketId = uri.searchParams.get('t') ?? '';
    expect(result.content[0].text).not.toContain(ticketId);
    expect(result.content[0].text).not.toContain(seen.uris[0]);
    expect(result.content[0].text).not.toContain(String(attachmentPort));
  });

  it('turns a Graph attachment path into markdown when Graph does state a truthful length', async () => {
    // The secondary fixture: Graph is not contractually forbidden from ever
    // sending content-length, so the route's guard must also pass a truthful
    // positive length through unchanged rather than only ever omitting it.
    stubDownloadStream(DOCUMENT_BYTES.length);

    const client = await agent();
    const result = (await client.callTool({
      name: 'read-document',
      arguments: { target: MAIL_ATTACHMENT },
    })) as { isError?: boolean; content: Array<{ text: string }> };

    expect(result.isError).toBeFalsy();
    expect(result.content[0].text).toBe(`# report.pdf\n\n${DOCUMENT_BYTES}\n`);
    expect(seen.statuses).toEqual([200]);

    // The declared length matches the real byte count exactly, and is never
    // the empty-body lie.
    expect(seen.contentLengths[0]).toBe(String(DOCUMENT_BYTES.length));
    expect(seen.contentLengths[0]).not.toBe('0');
  });

  it('does not register the byte tools on the same server', async () => {
    const client = await agent();
    const names = (await client.listTools()).tools.map((t) => t.name);
    expect(names).toContain('read-document');
    expect(names).not.toContain('download-bytes');
    expect(names).not.toContain('get-download-url');
    expect(names).not.toContain('get-mail-message-mime');
  });

  it('refuses a target the ticket grammar does not cover, without dialling anything', async () => {
    const client = await agent();
    const result = (await client.callTool({
      name: 'read-document',
      arguments: { target: '/me/drive/root/children' },
    })) as { content: Array<{ text: string }> };

    expect(JSON.parse(result.content[0].text).error).toBe('invalid_target');
    expect(seen.uris).toHaveLength(0);
  });
});
