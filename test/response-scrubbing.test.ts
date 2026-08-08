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
