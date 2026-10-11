/**
 * `$expand` is the third path to the same failure, and the one nobody saw for
 * eight days.
 *
 * Measured on one real message: `get-mail-message{select:"id,subject"}` returns
 * 246 bytes; the same call plus `expand:["attachments"]` returns 291,572. select
 * does not suppress contentBytes, 38 tools expose expand, and the shipped
 * description of the parameter says "e.g. attachments on a message or event" --
 * so this is the tool's documented behaviour rather than a misuse of it.
 *
 * Refused once, in executeGraphTool, which is where every registered tool AND
 * discovery's execute-tool both land. This is defence in depth, not the
 * invariant: the response scrubber is what holds "no bytes reach the model". It
 * exists because a refusal naming read-document teaches the model the right call,
 * where a stripped field only tells it something disappeared.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import {
  findByteInliningExpand,
  findByteInliningExpandInBatch,
  registerGraphTools,
} from '../src/graph-tools.js';
import {
  configureAttachmentProxy,
  resetAttachmentProxy,
} from '../src/lib/attachment-proxy-runtime.js';
import type { AttachmentProxyClient } from '../src/lib/attachment-proxy.js';
import type GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn(), verbose: vi.fn() },
  enableConsoleLogging: vi.fn(),
}));

describe('findByteInliningExpand', () => {
  it('matches every spelling a caller can reach', () => {
    expect(findByteInliningExpand({ expand: ['attachments'] })).toBe('attachments');
    expect(findByteInliningExpand({ $expand: ['attachments'] })).toBe('attachments');
    expect(findByteInliningExpand({ expand: 'attachments' })).toBe('attachments');
    expect(findByteInliningExpand({ expand: 'ATTACHMENTS' })).toBe('ATTACHMENTS');
    expect(findByteInliningExpand({ expand: ['organizer', 'attachments'] })).toBe('attachments');
    expect(findByteInliningExpand({ expand: 'organizer,attachments' })).toBe('attachments');
    expect(findByteInliningExpand({ expand: ['attachments($select=name)'] })).toBe(
      'attachments($select=name)'
    );
    expect(findByteInliningExpand({ expand: ['attachments/microsoft.graph.fileAttachment'] })).toBe(
      'attachments/microsoft.graph.fileAttachment'
    );
    // The same class, one entity over: chatMessage hostedContents inlines
    // contentBytes exactly the way message attachments do.
    expect(findByteInliningExpand({ expand: ['hostedContents'] })).toBe('hostedContents');
    // Leading/trailing whitespace around an entry, and around a comma-joined piece.
    expect(findByteInliningExpand({ expand: ['  attachments  '] })).toBe('attachments');
    expect(findByteInliningExpand({ expand: 'organizer, attachments' })).toBe('attachments');
  });

  it('leaves every other expand alone', () => {
    expect(findByteInliningExpand({ expand: ['organizer'] })).toBeNull();
    expect(findByteInliningExpand({ expand: ['singleValueExtendedProperties'] })).toBeNull();
    expect(findByteInliningExpand({ expand: [] })).toBeNull();
    expect(findByteInliningExpand({})).toBeNull();
    expect(findByteInliningExpand({ expand: null })).toBeNull();
    expect(findByteInliningExpand({ expand: 123 })).toBeNull();
    // A folder or a property that merely CONTAINS the word.
    expect(findByteInliningExpand({ expand: ['attachmentSessions'] })).toBeNull();
  });

  it('checks $expand and expand independently, so a present-but-empty one cannot mask the other', () => {
    // `.passthrough()` on every tool's input schema (and execute-tool's
    // `z.record(z.any())`) means a caller can hand both spellings at once.
    // `params.$expand ?? params.expand` would pick the empty array and never
    // look at `expand`, letting `['attachments']` through unseen.
    expect(findByteInliningExpand({ $expand: [], expand: ['attachments'] })).toBe('attachments');
    expect(findByteInliningExpand({ $expand: ['attachments'], expand: [] })).toBe('attachments');
    expect(findByteInliningExpand({ $expand: ['organizer'], expand: ['attachments'] })).toBe(
      'attachments'
    );
  });

  it("finds $expand nested inside another property's parenthesised options, at any depth", () => {
    // A real Graph pattern: a recurring event's expanded instances, each
    // expanding their own attachments. The outer head token is "instances",
    // so a check that only looks before the first "(" never sees it.
    expect(findByteInliningExpand({ expand: ['instances($expand=attachments)'] })).toBe(
      'attachments'
    );
    // Two levels of nesting -- detection must not be hardcoded to depth 1.
    expect(findByteInliningExpand({ expand: ['a($expand=b($expand=attachments))'] })).toBe(
      'attachments'
    );
    // A nested expand that is itself harmless stays harmless.
    expect(findByteInliningExpand({ expand: ['instances($expand=organizer)'] })).toBeNull();
  });

  it('matches the key the way the executor does: any case, with or without the $', () => {
    // The executor's query loop folds every one of these onto a real `$expand` that is sent to
    // Graph, so a guard reading only the literal `expand`/`$expand` let them inline bytes.
    for (const key of ['Expand', '$EXPAND', 'EXPAND', '$Expand', 'eXpAnD']) {
      expect(findByteInliningExpand({ [key]: ['attachments'] })).toBe('attachments');
      expect(findByteInliningExpand({ [key]: 'organizer,attachments' })).toBe('attachments');
    }
  });

  it('scans every matching key, so an empty one cannot mask a case-variant one', () => {
    expect(findByteInliningExpand({ $expand: [], EXPAND: ['attachments'] })).toBe('attachments');
    expect(findByteInliningExpand({ $expand: '', Expand: 'attachments' })).toBe('attachments');
    expect(findByteInliningExpand({ expand: ['organizer'], $Expand: ['attachments'] })).toBe(
      'attachments'
    );
    expect(findByteInliningExpand({ expand: [], $EXPAND: [] })).toBeNull();
  });

  it('reads a value the way the executor stringifies it', () => {
    expect(findByteInliningExpand({ expand: [['attachments']] })).toBe('attachments');
    expect(findByteInliningExpand({ expand: [' attachments'] })).toBe('attachments');
    expect(findByteInliningExpand({ expand: ['Attachments($select=name)', 'organizer'] })).toBe(
      'Attachments($select=name)'
    );
    expect(findByteInliningExpand({ expand: ['a($expand=b($expand=Attachments))'] })).toBe(
      'Attachments'
    );
    expect(findByteInliningExpand({ expand: 'organizer, hostedContents' })).toBe('hostedContents');
  });

  it('leaves an unrelated parameter alone even when it merely contains the word', () => {
    expect(findByteInliningExpand({ expandExtendedProperties: true })).toBeNull();
    expect(findByteInliningExpand({ notexpand: ['attachments'] })).toBeNull();
    expect(findByteInliningExpand({ select: 'attachments' })).toBeNull();
  });
});

describe('findByteInliningExpandInBatch', () => {
  const batch = (url: string) => ({ requests: [{ id: '1', method: 'GET', url }] });

  it('finds a byte-inlining expand in a sub-request URL, as an object or a JSON string', () => {
    expect(findByteInliningExpandInBatch(batch('/me/messages/AAA?$expand=attachments'))).toBe(
      'attachments'
    );
    expect(
      findByteInliningExpandInBatch(
        JSON.stringify(batch('/me/messages/AAA?$select=id&$expand=attachments($select=name)'))
      )
    ).toBe('attachments($select=name)');
  });

  it('matches the option name and value in any case, percent-encoding or nesting', () => {
    for (const url of [
      '/me/messages/AAA?$EXPAND=ATTACHMENTS',
      '/me/messages/AAA?expand=attachments',
      '/me/messages/AAA?%24expand=attachments',
      '/me/messages/AAA?%2524expand=attachments',
      '/me/messages/AAA?$expand=organizer%2Cattachments',
      '/me/events/AAA?$expand=instances($expand=attachments)',
      '/me/messages/AAA?$expand=%20attachments',
    ]) {
      expect(findByteInliningExpandInBatch(batch(url)), url).not.toBeNull();
    }
  });

  it('leaves harmless sub-requests alone', () => {
    expect(findByteInliningExpandInBatch(batch('/me/messages/AAA?$expand=organizer'))).toBeNull();
    expect(findByteInliningExpandInBatch(batch('/me/messages/AAA?$select=attachments'))).toBeNull();
    expect(findByteInliningExpandInBatch(batch('/me/messages/AAA'))).toBeNull();
    expect(findByteInliningExpandInBatch('not json')).toBeNull();
    expect(findByteInliningExpandInBatch(undefined)).toBeNull();
    expect(findByteInliningExpandInBatch({ requests: 'x' })).toBeNull();
  });
});

describe('the guard in executeGraphTool', () => {
  const graphClient = {
    graphRequest: vi.fn().mockResolvedValue({ content: [{ type: 'text', text: '{}' }] }),
  } as unknown as GraphClient;

  async function connect(proxyOn: boolean, tools = '^get-mail-message$'): Promise<Client> {
    if (proxyOn) {
      configureAttachmentProxy({
        client: {
          convertToMarkdown: async () => ({ ok: true, markdown: '' }),
        } as unknown as AttachmentProxyClient,
        url: 'http://docglean:8080/mcp',
      });
    } else {
      resetAttachmentProxy();
    }
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    registerGraphTools(server, graphClient, {
      enabledTools: tools,
      httpMode: true,
      attachmentProxy: proxyOn,
    });
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await server.connect(serverTransport);
    const client = new Client({ name: 'test-client', version: '1.0.0' });
    await client.connect(clientTransport);
    return client;
  }

  beforeEach(() => {
    vi.clearAllMocks();
  });

  afterEach(() => {
    resetAttachmentProxy();
  });

  it('refuses expand: ["attachments"] in proxy mode, and never dials Graph', async () => {
    const client = await connect(true);
    const result = (await client.callTool({
      name: 'get-mail-message',
      arguments: { messageId: 'AAA', select: 'id,subject', expand: ['attachments'] },
    })) as { isError?: boolean; content: Array<{ text: string }> };

    expect(result.isError).toBe(true);
    const body = JSON.parse(result.content[0].text);
    expect(body.error).toBe('expand_not_allowed');
    expect(body.expand).toBe('attachments');
    // The refusal has to name the alternative, or the model just tries again.
    expect(body.message).toContain('read-document');
    expect(graphClient.graphRequest).not.toHaveBeenCalled();
  });

  it.each(['Expand', '$EXPAND', 'EXPAND', '$Expand'])(
    'refuses a case-variant %s key in proxy mode, and never dials Graph',
    async (key) => {
      const client = await connect(true);
      const result = (await client.callTool({
        name: 'get-mail-message',
        arguments: { messageId: 'AAA', $expand: [], [key]: ['attachments'] },
      })) as { isError?: boolean; content: Array<{ text: string }> };

      expect(result.isError).toBe(true);
      expect(JSON.parse(result.content[0].text).error).toBe('expand_not_allowed');
      expect(graphClient.graphRequest).not.toHaveBeenCalled();
    }
  );

  it('refuses a graph-batch sub-request that expands attachments, and never dials Graph', async () => {
    const client = await connect(true, '^graph-batch$');
    const result = (await client.callTool({
      name: 'graph-batch',
      arguments: {
        body: {
          requests: [{ id: '1', method: 'GET', url: '/me/messages/AAA?$Expand=attachments' }],
        },
      },
    })) as { isError?: boolean; content: Array<{ text: string }> };

    expect(result.isError).toBe(true);
    expect(JSON.parse(result.content[0].text).error).toBe('expand_not_allowed');
    expect(graphClient.graphRequest).not.toHaveBeenCalled();
  });

  it('lets a harmless graph-batch through in proxy mode', async () => {
    const client = await connect(true, '^graph-batch$');
    const result = (await client.callTool({
      name: 'graph-batch',
      arguments: { body: { requests: [{ id: '1', method: 'GET', url: '/me/messages/AAA' }] } },
    })) as { isError?: boolean };

    expect(result.isError).toBeFalsy();
    expect(graphClient.graphRequest).toHaveBeenCalled();
  });

  it('allows expand: ["attachments"] when the flag is off', async () => {
    const client = await connect(false);
    const result = (await client.callTool({
      name: 'get-mail-message',
      arguments: { messageId: 'AAA', expand: ['attachments'] },
    })) as { isError?: boolean };

    expect(result.isError).toBeFalsy();
    expect(graphClient.graphRequest).toHaveBeenCalled();
  });

  it('leaves a harmless expand working in proxy mode', async () => {
    const client = await connect(true);
    const result = (await client.callTool({
      name: 'get-mail-message',
      arguments: { messageId: 'AAA', expand: ['singleValueExtendedProperties'] },
    })) as { isError?: boolean };

    expect(result.isError).toBeFalsy();
    expect(graphClient.graphRequest).toHaveBeenCalled();
  });
});
