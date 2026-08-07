import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { UTILITY_TOOLS } from '../src/graph-tools.js';
import { requestContext } from '../src/request-context.js';
import { AttachmentTicketStore } from '../src/lib/attachment-tickets.js';
import {
  configureAttachmentMinting,
  resetAttachmentMinting,
} from '../src/lib/attachment-minting.js';

/**
 * Regression cover for an authority escalation.
 *
 * The mint guard originally read `if (authManager?.isOAuthModeEnabled())`. That
 * predicate is true only for `MS365_MCP_OAUTH_TOKEN` and the oauth-provider
 * path -- it is **false** in plain `--http` bearer mode and in `--obo`, both of
 * which still run a tool inside a request context carrying the *caller's*
 * token. So in those modes `download-bytes` read as the caller while a redeemed
 * ticket read as whatever account the server had cached: mint under one
 * identity, fetch under another.
 *
 * These tests pin the corrected predicate. If either half is dropped again, the
 * "grants no authority the caller did not already have" claim in the README and
 * in `mintDownloadUrl`'s docstring stops being true, and one of these fails.
 */
describe('minting refuses whenever Graph identity comes from the request', () => {
  const tool = UTILITY_TOOLS.find((t) => t.name === 'get-download-url')!;
  const MAIL_ATTACHMENT = '/me/messages/AAA/attachments/BBB/$value';

  function ctx(overrides: Record<string, unknown> = {}) {
    return {
      graphClient: {} as never,
      authManager: {
        isOAuthModeEnabled: () => false,
        isMultiAccount: async () => false,
        getTokenForAccount: async () => 'SERVER_OWN_TOKEN',
        ...overrides,
      } as never,
      multiAccount: false,
      accountNames: [],
    };
  }

  function parse(result: { content: Array<{ text: string }> }) {
    return JSON.parse(result.content[0].text);
  }

  beforeEach(() => {
    configureAttachmentMinting({
      store: new AttachmentTicketStore(120),
      config: { base: 'http://m365:3000', key: 'k', keyId: '1', ttlSeconds: 120 },
    });
  });

  afterEach(() => resetAttachmentMinting());

  it('mints when identity is the server’s own cached token', async () => {
    const result = await tool.execute({ target: MAIL_ATTACHMENT }, ctx());
    expect(parse(result as never).downloadUrl).toMatch(/^http:\/\/m365:3000\/attachment\?/);
  });

  it('refuses in bearer/OBO mode, where a request token is present', async () => {
    const result = await requestContext.run({ accessToken: 'CALLER_TOKEN' }, () =>
      tool.execute({ target: MAIL_ATTACHMENT }, ctx())
    );
    expect(result.isError).toBe(true);
    const body = parse(result as never);
    expect(body.downloadUrl).toBeUndefined();
    expect(body.error).toMatch(/identity comes from the request/i);
  });

  it('refuses when MS365_MCP_OAUTH_TOKEN-style OAuth mode is on', async () => {
    const result = await tool.execute(
      { target: MAIL_ATTACHMENT },
      ctx({ isOAuthModeEnabled: () => true })
    );
    expect(result.isError).toBe(true);
    expect(parse(result as never).downloadUrl).toBeUndefined();
  });

  it('still refuses, rather than minting, for a non-byte Graph path', async () => {
    // The authority argument only holds for byte endpoints; an arbitrary Graph
    // path must stay refused even with minting enabled.
    const result = await tool.execute({ target: '/me/messages/AAA' }, ctx());
    expect(result.isError).toBe(true);
    expect(parse(result as never).downloadUrl).toBeUndefined();
  });

  it('does not mint at all when the feature is off', async () => {
    resetAttachmentMinting();
    const result = await tool.execute({ target: MAIL_ATTACHMENT }, ctx());
    expect(result.isError).toBe(true);
    expect(parse(result as never).error).toMatch(/do not expose a pre-authenticated/i);
  });
});

/**
 * The descriptions are what an LLM reads to decide what to call, so a description
 * that is wrong in either direction is a real failure: the agent either rules out
 * a call that would succeed on a flag-enabled server, or makes one that cannot
 * succeed without the flag. These pin the guidance to both deployment modes.
 */
describe('download guidance stays true in both deployment modes', () => {
  const utilityTexts: Array<{ tool: string; where: string; text: string }> = [];
  for (const tool of UTILITY_TOOLS) {
    // stdio-only tools are exempt from the flag rule below: minting needs HTTP,
    // and `--enable-attachment-urls` is warned about and ignored in stdio mode
    // (server.ts). A stdio-only tool's text can never be read in a deployment
    // where the flag does anything, so an unqualified claim there stays true.
    if (tool.stdioOnly) continue;
    utilityTexts.push({ tool: tool.name, where: 'description', text: tool.description });
    const schema = tool.buildSchema({
      graphClient: {} as never,
      authManager: undefined as never,
      multiAccount: true,
      accountNames: [],
    } as never);
    for (const [field, zodType] of Object.entries(schema)) {
      utilityTexts.push({
        tool: tool.name,
        where: `${field} schema`,
        text: zodType.description ?? '',
      });
    }
  }

  // Class rule, not a spot check: any tool text that routes mail/event
  // attachments, recordings, or other /$value byte endpoints at get-download-url
  // -- including get-download-url's own text -- must name the flag those targets
  // depend on. Without it the text is false in one of the two deployments.
  it('names --enable-attachment-urls wherever byte endpoints are tied to get-download-url', () => {
    const BYTE_ENDPOINT = /attachment|recording|\$value/i;
    const offenders = utilityTexts
      .filter(({ tool, text }) => tool === 'get-download-url' || text.includes('get-download-url'))
      .filter(({ text }) => BYTE_ENDPOINT.test(text))
      .filter(({ text }) => !text.includes('--enable-attachment-urls'))
      .map(({ tool, where }) => `${tool} ${where}`);

    expect(offenders).toEqual([]);
  });

  it('get-download-url states the unconditional drive/SharePoint case', () => {
    const tool = UTILITY_TOOLS.find((t) => t.name === 'get-download-url')!;
    expect(tool.description).toContain('always available for drive/SharePoint file content');
    expect(tool.description).toContain('@microsoft.graph.downloadUrl');
    expect(tool.description).toContain('needs no flag');
  });

  it('get-download-url states the minted case and its refusal conditions', () => {
    const tool = UTILITY_TOOLS.find((t) => t.name === 'get-download-url')!;
    expect(tool.description).toContain('--enable-attachment-urls');
    expect(tool.description).toContain('HTTP mode only');
    expect(tool.description).toContain('singleUse: true');
    // The identity guard that attachment-mint-identity pins above.
    expect(tool.description).toMatch(/OAuth, OBO, or bearer mode/);
    expect(tool.description).toMatch(/refused/i);
    expect(tool.description).toContain('download-bytes');
  });

  it('download-bytes keeps the drive/SharePoint advice and qualifies the rest', () => {
    const tool = UTILITY_TOOLS.find((t) => t.name === 'download-bytes')!;
    expect(tool.description).toContain('prefer get-download-url');
    expect(tool.description).toContain('--enable-attachment-urls');
  });
});
