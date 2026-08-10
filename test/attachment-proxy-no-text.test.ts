/**
 * A conversion that yields no text is an answer, not silence.
 *
 * The proxy can succeed and still have nothing to hand back: a scanned PDF has
 * no text layer, and a genuinely blank document has no text either. Both
 * arrive as `isError: false` with `markdown: ""`, and an empty string is the
 * one response a model cannot act on -- indistinguishable from a bug, so it
 * retries, varies `pages`, or reports the document as blank. The converter
 * says which case it is in a `content_status` field; this suite pins that the
 * field survives the wire derivation and reaches the caller as a stated
 * reason.
 *
 * The wire cases use a REAL `AttachmentProxyClient` over a `fetch` mock, not a
 * stub of the client's own interface: `content_status` has to survive
 * `interpretJsonRpcMessage`, and a stub at the client's interface hands back
 * whatever the test author wrote and could never prove that.
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
import { AttachmentProxyClient, interpretJsonRpcMessage } from '../src/lib/attachment-proxy.js';
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

function fakeGraphClient(): GraphClient {
  return {
    makeRequest: vi.fn(async () => ({
      name: 'scan.pdf',
      contentType: 'application/pdf',
      size: 15_728_640,
    })),
  } as unknown as GraphClient;
}

function jsonRpcResponse(message: unknown): Response {
  return new Response(JSON.stringify(message), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
}

/**
 * A real client whose only fake is `fetch`.
 *
 * The payload may be built from the `uri` the client was actually handed, so a
 * test can echo the live minted ticket URL back the way a chatty converter
 * would -- the only way to exercise redaction against the real ticket rather
 * than a string the test invented.
 */
function realProxyReturning(
  payload: Record<string, unknown> | ((requestedUri: string) => Record<string, unknown>)
): AttachmentProxyClient {
  const fetchImpl = (async (_input: unknown, init?: unknown) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body = JSON.parse(String((init as any)?.body ?? '{}'));
    const uri = body?.params?.arguments?.uri as string;
    return jsonRpcResponse({
      jsonrpc: '2.0',
      id: 1,
      result: {
        structuredContent: typeof payload === 'function' ? payload(uri) : payload,
      },
    });
  }) as unknown as typeof fetch;
  return new AttachmentProxyClient({ url: 'http://docglean:8080/mcp', fetchImpl });
}

describe('a conversion that produced no text', () => {
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

  async function read(
    payload: Record<string, unknown> | ((requestedUri: string) => Record<string, unknown>)
  ): Promise<{
    isError: boolean;
    text: string;
    body: Record<string, unknown>;
  }> {
    configureAttachmentProxy({
      client: realProxyReturning(payload),
      url: 'http://docglean:8080/mcp',
    });
    const result = (await (
      await connect()
    ).callTool({ name: 'read-document', arguments: { target: MAIL_ATTACHMENT } })) as {
      isError?: boolean;
      content: Array<{ text: string }>;
    };
    const text = result.content[0].text;
    let body: Record<string, unknown> = {};
    try {
      body = JSON.parse(text) as Record<string, unknown>;
    } catch {
      body = {};
    }
    return { isError: Boolean(result.isError), text, body };
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

  it('carries content_status through the wire derivation', () => {
    const outcome = interpretJsonRpcMessage(
      {
        jsonrpc: '2.0',
        id: 1,
        result: { structuredContent: { markdown: '', content_status: 'no_text_layer' } },
      },
      200
    );

    expect(outcome).toEqual({ ok: true, markdown: '', contentStatus: 'no_text_layer' });
  });

  it('omits contentStatus when the converter did not send one', () => {
    const outcome = interpretJsonRpcMessage(
      { jsonrpc: '2.0', id: 1, result: { structuredContent: { markdown: '# Body' } } },
      200
    );

    expect(outcome).toEqual({ ok: true, markdown: '# Body' });
  });

  it('reports a scan as a scan rather than returning an empty string', async () => {
    const result = await read({
      markdown: '',
      content_status: 'no_text_layer',
      pdf_classification: { pdf_type: 'scanned', pages_needing_ocr: [1, 2, 3] },
    });

    expect(result.isError).toBe(true);
    expect(result.body.error).toBe('no_text_content');
    expect(result.body.contentStatus).toBe('no_text_layer');
    expect(String(result.body.message)).toMatch(/no text layer/i);
    // The facts already probed for the mint, so the agent can describe the
    // thing it could not read without a second call.
    expect(result.body.name).toBe('scan.pdf');
    expect(result.body.contentType).toBe('application/pdf');
    expect(result.body.size).toBe(15_728_640);
  });

  it('distinguishes a blank document from a scan', async () => {
    const result = await read({ markdown: '', content_status: 'empty_document' });

    expect(result.isError).toBe(true);
    expect(result.body.error).toBe('no_text_content');
    expect(result.body.contentStatus).toBe('empty_document');
    expect(String(result.body.message)).toMatch(/no text/i);
    expect(String(result.body.message)).not.toMatch(/no text layer/i);
  });

  it('surfaces a status it has no wording for, verbatim', async () => {
    // The case this exists for: whatever a future converter reports -- an OCR
    // pass that ran and failed, say -- reaches the agent as a stated reason
    // instead of an empty string, with no edit here.
    const result = await read({ markdown: '', content_status: 'ocr_failed' });

    expect(result.isError).toBe(true);
    expect(result.body.error).toBe('no_text_content');
    expect(result.body.contentStatus).toBe('ocr_failed');
    expect(String(result.body.message)).toContain('ocr_failed');
  });

  it('still states a reason when the converter gave none', async () => {
    const result = await read({ markdown: '' });

    expect(result.isError).toBe(true);
    expect(result.body.error).toBe('no_text_content');
    expect(result.body.contentStatus).toBeUndefined();
    expect(String(result.body.message)).toMatch(/no reason/i);
  });

  it('treats whitespace-only markdown as no text', async () => {
    const result = await read({ markdown: '   \n\t  ', content_status: 'no_text_layer' });

    expect(result.isError).toBe(true);
    expect(result.body.error).toBe('no_text_content');
  });

  it('does not hijack a successful read that also carries a status', async () => {
    const result = await read({ markdown: '# Q3 report\n\nBody.', content_status: 'ok' });

    expect(result.isError).toBe(false);
    expect(result.text).toBe('# Q3 report\n\nBody.');
  });

  it('redacts a ticket the converter echoed into its status', async () => {
    // The LIVE minted URL, echoed back the way a chatty converter would, so the
    // assertion is about redaction rather than about a string this test invented.
    let echoed = '';
    const result = await read((requestedUri) => {
      echoed = requestedUri;
      return { markdown: '', content_status: `no_text_layer while fetching ${requestedUri}` };
    });

    expect(result.isError).toBe(true);
    expect(echoed).toContain('/attachment');
    const ticketId = new URL(echoed).searchParams.get('t') ?? '';
    expect(ticketId).not.toBe('');
    expect(result.text).not.toContain(ticketId);
    expect(result.text).not.toContain(echoed);
  });
});
