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
import type {
  AttachmentProxyClient,
  ConvertRequest,
  ConvertResult,
} from '../src/lib/attachment-proxy.js';
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
});
