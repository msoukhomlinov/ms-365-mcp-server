/**
 * The invariant's enforcement point.
 *
 * Banning download-bytes, then get-mail-message-mime, then expand patches the
 * three paths we know about. With 38 tools exposing expand, and upstream free to
 * add a byte-returning tool at any time, per-tool bans cannot hold "no tool on
 * this server returns raw bytes". A wrapper on tools/call can, including for a
 * tool that does not exist yet.
 *
 * read-document is the one exemption, and it is safe by construction: it returns
 * markdown the proxy produced, never bytes. Without the exemption a document
 * whose markdown happened to contain a long base64 block -- a code fence in a
 * technical PDF is enough -- would be silently mangled.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { z } from 'zod';
import { installResponseScrubbing } from '../src/response-scrubbing.js';
import logger from '../src/logger.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn(), verbose: vi.fn() },
  enableConsoleLogging: vi.fn(),
}));

/** 200 KB of valid base64 -- the shape a real 195 KB PDF attachment arrives in. */
const BIG_BASE64 = Buffer.from('A'.repeat(150_000)).toString('base64');
/** A Graph id in this deployment is ~152 characters. Must never be touched. */
const GRAPH_ID =
  'AAMkAGI2NDlhZTJmLTBhOTQtNDk0NC1hMTM4LTgxYmY1ZjZmM2QwMwBGAAAAAAB1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghij';

describe('installResponseScrubbing', () => {
  async function connect(): Promise<Client> {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    server.tool(
      'get-mail-message',
      'returns a message',
      { id: z.string().optional() },
      async () => ({
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({
              id: GRAPH_ID,
              subject: 'Q3',
              attachments: [{ name: 'report.pdf', contentBytes: BIG_BASE64 }],
            }),
          },
        ],
      })
    );
    server.tool(
      'read-document',
      'returns markdown',
      { target: z.string().optional() },
      async () => ({
        content: [{ type: 'text' as const, text: `# Report\n\n\`\`\`\n${BIG_BASE64}\n\`\`\`\n` }],
      })
    );
    server.tool('raw-text-tool', 'returns bare text', {}, async () => ({
      content: [{ type: 'text' as const, text: BIG_BASE64 }],
    }));

    installResponseScrubbing(server);

    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await server.connect(serverTransport);
    const client = new Client({ name: 'test-client', version: '1.0.0' });
    await client.connect(clientTransport);
    return client;
  }

  async function text(client: Client, name: string): Promise<string> {
    const result = (await client.callTool({ name, arguments: {} })) as {
      content: Array<{ text: string }>;
    };
    return result.content[0].text;
  }

  beforeEach(() => {
    vi.clearAllMocks();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("strips another tool's contentBytes while leaving the Graph id alone", async () => {
    const client = await connect();
    const body = JSON.parse(await text(client, 'get-mail-message'));

    expect(body.attachments[0].contentBytes).not.toBe(BIG_BASE64);
    expect(String(body.attachments[0].contentBytes)).toContain('stripped');
    // Identifiers are never touched -- that is what the 4,096 floor is for.
    expect(body.id).toBe(GRAPH_ID);
    expect(body.subject).toBe('Q3');
    expect(body.attachments[0].name).toBe('report.pdf');
  });

  it('leaves read-document output exactly as the proxy produced it', async () => {
    const client = await connect();
    const markdown = await text(client, 'read-document');
    expect(markdown).toContain(BIG_BASE64);
  });

  it('covers a bare non-JSON text body too', async () => {
    // The shape with no JSON envelope to hide in. get-mail-message-mime is the
    // shipped example, and it is unregistered in proxy mode -- but the scrubber
    // is the guard for the tool nobody has written yet, so it must not have a
    // hole shaped like "the response was not JSON".
    const client = await connect();
    const out = await text(client, 'raw-text-tool');
    expect(out).not.toBe(BIG_BASE64);
    expect(out).toContain('stripped');
  });

  it('names the field it stripped in a warn line', async () => {
    const client = await connect();
    await text(client, 'get-mail-message');
    const warned = (logger.warn as ReturnType<typeof vi.fn>).mock.calls.map((c) => String(c[0]));
    const line = warned.find((l) => l.includes('contentBytes'));
    // This is how a fourth path gets discovered -- in a log line, not in a
    // context blowout.
    expect(line, 'the scrubber must report the field name it stripped').toBeDefined();
    expect(line).toContain('get-mail-message');
  });

  it('refuses to install silently when the handler is missing', () => {
    // Unlike the $ref normaliser it is modelled on, this one carries an
    // invariant. A no-op fallback would be a server that believes it is
    // scrubbing and is not.
    const bare = new McpServer({ name: 'test', version: '1.0.0' });
    expect(() => installResponseScrubbing(bare)).toThrow(/tools\/call/);
  });
});

describe('installResponseScrubbing -- unconventional content shapes', () => {
  // The real low-level `Server` (what `server.server` actually is) wraps its
  // *own* `tools/call` handler registration in a second layer: it re-validates
  // whatever that handler returns against CallToolResultSchema, which requires
  // `content` to be an array, before the result goes anywhere. That is useful
  // defense in depth in this deployment's exact configuration, but it means a
  // handler retrieved from a real, fully-wired `McpServer` can't tell us
  // whether *this module's own logic* handled an unconventional shape, or
  // whether an unrelated downstream check simply rejected the whole call
  // first. A double that stands in for the low-level server -- a
  // `_requestHandlers` map plus a `setRequestHandler` that just writes into
  // it, no schema validation on either side -- isolates exactly what
  // `installResponseScrubbing` itself does with the `original` handler's
  // result, which is the actual thing under test here.
  type RawHandler = (request: unknown, extra: unknown) => Promise<unknown>;

  function installOverOriginal(original: RawHandler): RawHandler {
    const handlers = new Map<string, RawHandler>();
    handlers.set('tools/call', original);
    const fakeLowLevelServer = {
      _requestHandlers: handlers,
      setRequestHandler: (_schema: unknown, handler: RawHandler) => {
        handlers.set('tools/call', handler);
      },
    };
    const fakeServer = { server: fakeLowLevelServer } as unknown as McpServer;
    installResponseScrubbing(fakeServer);
    return handlers.get('tools/call')!;
  }

  const fakeRequest = { params: { name: 'weird-tool' } };

  it('scrubs a bare-string content body instead of iterating its characters', async () => {
    // Pre-fix: `for (const item of result.content ?? [])` over a string
    // iterates its *characters*. None of them is `{ type: 'text' }`, so the
    // loop does nothing and the payload survives untouched -- silently, with
    // no log line. This is the bug under review. Post-fix, `content` was never
    // a valid array, so the scrubbed value is wrapped in one -- the transport
    // still requires that shape regardless of what this module does.
    const handler = installOverOriginal(async () => ({ content: BIG_BASE64 }));
    const result = (await handler(fakeRequest, {})) as { content: Array<{ text: string }> };
    expect(Array.isArray(result.content)).toBe(true);
    expect(JSON.stringify(result.content)).not.toContain(BIG_BASE64);
    expect(result.content[0].text).toContain('stripped');
  });

  it('scrubs content given as a single object instead of an array of blocks', async () => {
    // Pre-fix: `for...of` over a non-array plain object throws (objects are
    // not iterable), so this shape fails a different way than the bare string
    // does, but it is still not scrubbed.
    const handler = installOverOriginal(async () => ({ content: { data: BIG_BASE64 } }));
    const result = (await handler(fakeRequest, {})) as { content: Array<{ text: string }> };
    expect(Array.isArray(result.content)).toBe(true);
    expect(JSON.stringify(result.content)).not.toContain(BIG_BASE64);
    expect(result.content[0].text).toContain('stripped');
  });

  it('scrubs a non-text content block (an image block with base64 in `data`)', async () => {
    // Pre-fix: `if (item?.type !== 'text' ...) continue;` skips this block
    // whole -- an image/audio/resource block's bytes live in `data` or
    // `resource.blob`, never `text`, so it walked straight past the only
    // per-item rule that existed. Post-fix, the block is converted to a text
    // block: a marker is not valid base64, and an image block's `data` must
    // be, so leaving it in place would make the block violate its own schema.
    const handler = installOverOriginal(async () => ({
      content: [{ type: 'image', data: BIG_BASE64, mimeType: 'image/png' }],
    }));
    const result = (await handler(fakeRequest, {})) as { content: Array<Record<string, unknown>> };
    const serialized = JSON.stringify(result.content);
    expect(serialized).not.toContain(BIG_BASE64);
    expect(serialized).toContain('stripped');
  });

  it('does not throw when content is absent, or the whole result is not an object', async () => {
    await expect(installOverOriginal(async () => ({}))(fakeRequest, {})).resolves.toEqual({});
    await expect(
      installOverOriginal(async () => undefined)(fakeRequest, {})
    ).resolves.toBeUndefined();
    await expect(installOverOriginal(async () => 'just a string')(fakeRequest, {})).resolves.toBe(
      'just a string'
    );
  });
});
