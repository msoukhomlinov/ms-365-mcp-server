import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerGraphTools } from '../src/graph-tools.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
  },
}));

vi.mock('../src/cloud-config.js', () => ({
  getCloudEndpoints: () => ({
    graphApi: 'https://graph.microsoft.com',
    authority: 'https://login.microsoftonline.com',
  }),
}));

vi.mock('../src/lib/microsoft-auth.js', () => ({
  refreshAccessToken: vi.fn(),
}));

const { default: GraphClient } = await import('../src/graph-client.js');

const mockAuthManager = {
  getToken: vi.fn().mockResolvedValue('mock-token'),
};

const mockSecrets = {
  clientId: 'test-client-id',
  tenantId: 'test-tenant-id',
  clientSecret: 'test-client-secret',
  cloudType: 'global' as const,
};

type Handler = (params: Record<string, unknown>) => Promise<{ content: { text: string }[] }>;

/**
 * Registers the real Graph tools (real generated endpoints, real endpoints.json)
 * against a real GraphClient and returns the handler for one tool.
 *
 * The transport (`fetch`) is the only thing mocked. GraphClient must stay real:
 * it is `formatJsonResponse`/`removeODataProps` that decides which `@odata.*`
 * fields ever reach the fetchAllPages merge, so a mocked client can hand the
 * aggregator a payload the real one could never produce.
 */
function handlerFor(alias: string, outputFormat: 'json' | 'toon' = 'json'): Handler {
  const server = new McpServer({ name: 'test', version: '1.0.0' });
  const handlers = new Map<string, Handler>();
  vi.spyOn(server, 'registerTool').mockImplementation(((
    name: string,
    _config: unknown,
    handler: Handler
  ) => {
    handlers.set(name, handler);
  }) as never);
  vi.spyOn(server, 'tool').mockImplementation((() => {}) as never);

  const graphClient = new GraphClient(mockAuthManager as never, mockSecrets, outputFormat);
  registerGraphTools(server, graphClient, false);

  const handler = handlers.get(alias);
  if (!handler) throw new Error(`Tool ${alias} was not registered`);
  return handler;
}

/** Queues one JSON body per fetch call, in order. */
function mockPages(...bodies: unknown[]) {
  const queue = [...bodies];
  return vi.spyOn(globalThis, 'fetch').mockImplementation(async () => {
    const body = queue.length > 0 ? queue.shift() : { value: [] };
    return new Response(JSON.stringify(body), {
      status: 200,
      headers: { 'content-type': 'application/json' },
    });
  });
}

const NEXT_1 = 'https://graph.microsoft.com/v1.0/me/messages?$skiptoken=page2';
const NEXT_2 = 'https://graph.microsoft.com/v1.0/me/messages?$skiptoken=page3';

async function callAndParse(handler: Handler, params: Record<string, unknown>) {
  const result = await handler(params);
  return JSON.parse(result.content[0].text) as Record<string, unknown>;
}

describe('fetchAllPages and @odata.count', () => {
  let mockFetch: ReturnType<typeof mockPages> | undefined;
  const prevMaxPages = process.env.MS365_MCP_MAX_PAGES;

  beforeEach(() => {
    vi.clearAllMocks();
    mockAuthManager.getToken.mockResolvedValue('mock-token');
  });

  afterEach(() => {
    mockFetch?.mockRestore();
    mockFetch = undefined;
    if (prevMaxPages === undefined) delete process.env.MS365_MCP_MAX_PAGES;
    else process.env.MS365_MCP_MAX_PAGES = prevMaxPages;
  });

  it('forwards $count=true to Graph', async () => {
    mockFetch = mockPages({ '@odata.count': 1, value: [{ id: '1' }] });

    await callAndParse(handlerFor('list-mail-messages'), { count: true });

    const url = String(mockFetch.mock.calls[0][0]);
    expect(url).toContain('count=true');
  });

  it('keeps @odata.count on a single-page response the caller asked to count', async () => {
    mockFetch = mockPages({
      '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#messages',
      '@odata.count': 2,
      value: [{ id: '1' }, { id: '2' }],
    });

    const parsed = await callAndParse(handlerFor('list-mail-messages'), {
      count: true,
      fetchAllPages: true,
    });

    expect(parsed['@odata.count']).toBe(2);
    expect(parsed['@odata.context']).toBeUndefined();
    expect(parsed['@odata.nextLink']).toBeUndefined();
    expect(mockFetch).toHaveBeenCalledTimes(1);
  });

  it('keeps a legitimate zero count instead of dropping it as falsy', async () => {
    mockFetch = mockPages({ '@odata.count': 0, value: [] });

    const parsed = await callAndParse(handlerFor('list-mail-messages'), {
      count: true,
      fetchAllPages: true,
    });

    expect(parsed['@odata.count']).toBe(0);
    expect(parsed.value).toEqual([]);
  });

  it('passes Graph’s count through unchanged when it differs from the aggregate', async () => {
    // $count describes the filtered collection, not the items in `value`, so the
    // merge must not "correct" it to the item tally. Every non-fetchAllPages
    // response already reports a count larger than its own page.
    mockFetch = mockPages(
      { '@odata.count': 2, '@odata.nextLink': NEXT_1, value: [{ id: '1' }, { id: '2' }] },
      { '@odata.count': 2, value: [{ id: '3' }] }
    );

    const parsed = await callAndParse(handlerFor('list-mail-messages'), {
      count: true,
      fetchAllPages: true,
    });

    expect((parsed.value as unknown[]).map((v) => (v as { id: string }).id)).toEqual([
      '1',
      '2',
      '3',
    ]);
    expect(parsed['@odata.count']).toBe(2);
    expect(parsed['@odata.nextLink']).toBeUndefined();
    expect(mockFetch).toHaveBeenCalledTimes(2);
  });

  it('passes a null count through rather than inventing an aggregate', async () => {
    mockFetch = mockPages(
      { '@odata.count': null, '@odata.nextLink': NEXT_1, value: [{ id: '1' }] },
      { value: [{ id: '2' }] }
    );

    const parsed = await callAndParse(handlerFor('list-mail-messages'), {
      count: true,
      fetchAllPages: true,
    });

    expect(parsed['@odata.count']).toBeNull();
  });

  it('keeps the pre-skip count when the request started after a $skip', async () => {
    // $skip=20 over a 100-item collection: Graph reports 100, the merge collects
    // the 80 remaining items. Rewriting the count to the tally would destroy the
    // collection size the caller asked for.
    mockFetch = mockPages(
      { '@odata.count': 100, '@odata.nextLink': NEXT_1, value: [{ id: '1' }, { id: '2' }] },
      { '@odata.count': 100, value: [{ id: '3' }] }
    );

    const parsed = await callAndParse(handlerFor('list-mail-messages'), {
      count: true,
      skip: 20,
      fetchAllPages: true,
    });

    expect(String(mockFetch.mock.calls[0][0])).toContain('skip=20');
    expect((parsed.value as unknown[]).length).toBe(3);
    expect(parsed['@odata.count']).toBe(100);
  });

  it('does not invent a count the caller never requested', async () => {
    mockFetch = mockPages(
      { '@odata.nextLink': NEXT_1, value: [{ id: '1' }] },
      { value: [{ id: '2' }] }
    );

    const parsed = await callAndParse(handlerFor('list-mail-messages'), { fetchAllPages: true });

    expect('@odata.count' in parsed).toBe(false);
    expect((parsed.value as unknown[]).length).toBe(2);
  });

  it('keeps the live nextLink and the server count when pagination is truncated', async () => {
    // maxPages=2 stops the loop with page 3 still outstanding. Deleting the
    // nextLink here would tell the model the collection was fully fetched, and
    // rewriting the count to 2 would erase the true total.
    process.env.MS365_MCP_MAX_PAGES = '2';
    mockFetch = mockPages(
      { '@odata.count': 3, '@odata.nextLink': NEXT_1, value: [{ id: '1' }] },
      { '@odata.count': 3, '@odata.nextLink': NEXT_2, value: [{ id: '2' }] },
      { '@odata.count': 3, value: [{ id: '3' }] }
    );

    const parsed = await callAndParse(handlerFor('list-mail-messages'), {
      count: true,
      fetchAllPages: true,
    });

    expect(mockFetch).toHaveBeenCalledTimes(2);
    expect((parsed.value as unknown[]).length).toBe(2);
    // The most recent nextLink, not page one's already-consumed token.
    expect(parsed['@odata.nextLink']).toBe(NEXT_2);
    expect(parsed['@odata.count']).toBe(3);
  });

  it('still drops the nextLink when aggregation actually completes', async () => {
    mockFetch = mockPages(
      { '@odata.nextLink': NEXT_1, value: [{ id: '1' }] },
      { '@odata.nextLink': NEXT_2, value: [{ id: '2' }] },
      { value: [{ id: '3' }] }
    );

    const parsed = await callAndParse(handlerFor('list-mail-messages'), { fetchAllPages: true });

    expect(mockFetch).toHaveBeenCalledTimes(3);
    expect(parsed['@odata.nextLink']).toBeUndefined();
  });
});
