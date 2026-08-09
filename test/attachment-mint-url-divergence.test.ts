/**
 * Codex P1 (PR #16, reviewed commit 5da1c847): the mint gate validated a
 * STRING (`target.endsWith('/$value')`) while what actually matters is the
 * PATHNAME of the URL `GraphClient.performRequest` builds.
 *
 * `performRequest` (src/graph-client.ts) concatenates `target` straight onto
 * the Graph base URL and hands the result to `fetch()`, which parses it per
 * the WHATWG URL standard. That parser treats several raw characters as
 * structural rather than literal:
 *   - `#` starts a fragment, which is never sent to the server at all.
 *   - `?` starts a query string, which Graph's path routing ignores.
 *   - `\` is folded into `/` for "special" schemes, https included.
 *   - `.`/`..` path segments are resolved away before the request is sent.
 *   - C0 controls are stripped or otherwise mangled.
 * A target crafted so the raw string ends in `/$value` (satisfying the old
 * check) can still make the wire request name a completely different
 * resource -- for a mail/event attachment, its metadata endpoint, whose JSON
 * body carries `contentBytes` and is streamed with no scrubber by the
 * redemption route in attachment-route.ts. This file pins the fix: the mint
 * gate must reject every one of the above, everywhere it is enforced,
 * without rejecting the legitimate targets it exists to allow through.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { isMintableTarget, UTILITY_TOOLS } from '../src/graph-tools.js';
import { AttachmentTicketStore } from '../src/lib/attachment-tickets.js';
import {
  configureAttachmentMinting,
  resetAttachmentMinting,
} from '../src/lib/attachment-minting.js';
import {
  configureAttachmentProxy,
  resetAttachmentProxy,
} from '../src/lib/attachment-proxy-runtime.js';
import type { AttachmentProxyClient, ConvertResult } from '../src/lib/attachment-proxy.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn(), verbose: vi.fn() },
  enableConsoleLogging: vi.fn(),
}));

const LEGIT_MAIL_ATTACHMENT = '/me/messages/AAA/attachments/BBB/$value';

/** Every raw-string evasion this gate must reject, with the divergent pathname it names. */
const EVASIONS: Array<{ label: string; target: string }> = [
  {
    label: 'a literal # fragment (the exact Codex PoC)',
    target: '/me/messages/M/attachments/A#/$value',
  },
  {
    label: 'a literal ? query string',
    target: '/me/messages/M/attachments/A?x=/$value',
  },
  {
    label: 'a percent-encoded %23 (encoded #)',
    target: '/me/messages/M/attachments/A%23/$value',
  },
  {
    label: 'a percent-encoded %3F (encoded ?)',
    target: '/me/messages/M/attachments/A%3Fx=/$value',
  },
  {
    // A backslash is folded into '/' by the WHATWG URL parser for "special"
    // schemes (https included), so a raw backslash mid-path can silently
    // reshape the requested pathname while the raw string still ends in the
    // literal suffix `endsWith('/$value')` checks for.
    label: 'a backslash path separator',
    target: '/me/messages/M/attachments/A\\foo/$value',
  },
  {
    // The raw string matches MAIL_EVENT_ATTACHMENT_TARGET (its regex only
    // anchors the `/me/messages/{id}/attachments/` prefix) and ends in
    // `/$value`, so the old string-only check accepted it -- but the `..`
    // segments are resolved away before the request is sent, so the ticket
    // this mints actually redeems attachment A2, not the validated A.
    label: 'a .. path-traversal segment that resolves to a different attachment',
    target: '/me/messages/M/attachments/A/../../attachments/A2/$value',
  },
  {
    label: 'a bare . path segment',
    target: '/me/messages/M/attachments/A/./$value',
  },
  {
    label: 'an embedded control character',
    target: '/me/messages/M/attachments/A\t/$value',
  },
];

describe('isMintableTarget rejects targets whose validated string diverges from the requested pathname', () => {
  for (const { label, target } of EVASIONS) {
    it(`rejects ${label}`, () => {
      expect(isMintableTarget(target)).toBe(false);
    });
  }

  it('still mints a legitimate mail attachment target', () => {
    expect(isMintableTarget(LEGIT_MAIL_ATTACHMENT)).toBe(true);
  });

  it('still mints a legitimate event attachment target', () => {
    expect(isMintableTarget('/me/events/EEE/attachments/BBB/$value')).toBe(true);
  });

  it('still mints a legitimate drive/SharePoint content target', () => {
    expect(isMintableTarget('/drives/DRIVE1/items/ITEM1/content')).toBe(true);
  });

  it('still mints a legitimate drive root:/path content target with spaces and unicode', () => {
    expect(isMintableTarget('/me/drive/root:/My Folder/Fïle (1).docx:/content')).toBe(true);
  });

  it('still mints a legitimate meeting recording target', () => {
    expect(isMintableTarget('/me/onlineMeetings/MEET1/recordings/REC1/content')).toBe(true);
  });

  it('still mints a legitimate generic /$value endpoint', () => {
    expect(isMintableTarget('/me/photo/$value')).toBe(true);
  });

  it('still refuses a mail attachment target missing the /$value suffix (unchanged behavior)', () => {
    expect(isMintableTarget('/me/messages/AAA/attachments/BBB')).toBe(false);
  });
});

/**
 * Codex P1 #2 (PR #16, src/graph-tools.ts:539): the first fix's denylist
 * caught literal `..` but missed the URL Standard's percent-encoded
 * dot-segment spellings. `/me/messages/M/attachments/A/%2e%2e/%2e%2e/attachments/A2/$value`
 * passed the first fix's `hasUrlDivergentSyntax` and `isMintableTarget`
 * unchanged, because neither treated `%2e%2e` as a `..`. Verified locally
 * (Node, simulating `performRequest`'s own concatenation):
 *
 *   new URL('https://graph.microsoft.com/v1.0' +
 *     '/me/messages/M/attachments/A/%2e%2e/%2e%2e/attachments/A2/$value').pathname
 *   -> '/v1.0/me/messages/M/attachments/A2/$value'
 *
 * i.e. the ticket this mints redeems attachment A2, never A -- silently
 * substituting a different attachment than the one the caller named. The
 * fix is architectural, not another denylist entry: `isMintableTarget` now
 * runs its family/suffix grammar against the RESOLVED pathname (built the
 * same way `performRequest` builds it) rather than the raw string, and
 * separately rejects any raw path segment that is a single-/double-dot
 * segment per the URL Standard's own closed definition (`.`, `%2e`, `..`,
 * `.%2e`, `%2e.`, `%2e%2e`, case-insensitively) -- not a guessed encoding
 * list, but the spec's own enumeration.
 */
describe('isMintableTarget rejects percent-encoded dot-segment spellings (Codex P1 #2)', () => {
  const DOT_SEGMENT_EVASIONS: Array<{ label: string; target: string }> = [
    {
      label: 'the full Codex payload verbatim',
      target: '/me/messages/M/attachments/A/%2e%2e/%2e%2e/attachments/A2/$value',
    },
    {
      label: 'a single %2e%2e segment',
      target: '/me/messages/M/attachments/A/%2e%2e/attachments/A2/$value',
    },
    {
      label: 'mixed-case %2E%2E',
      target: '/me/messages/M/attachments/A/%2E%2E/%2E%2E/attachments/A2/$value',
    },
    {
      label: 'the mixed literal-dot/encoded-dot spelling .%2e',
      target: '/me/messages/M/attachments/A/.%2e/.%2e/attachments/A2/$value',
    },
    {
      label: 'the mixed encoded-dot/literal-dot spelling %2e.',
      target: '/me/messages/M/attachments/A/%2e./%2e./attachments/A2/$value',
    },
    {
      label: 'a single-dot %2e segment (no traversal, still a non-literal path segment)',
      target: '/me/messages/M/attachments/A/%2e/$value',
    },
  ];

  for (const { label, target } of DOT_SEGMENT_EVASIONS) {
    it(`rejects ${label}`, () => {
      expect(isMintableTarget(target)).toBe(false);
    });
  }

  it('proves the rejected payload really does resolve to a different attachment (the real damage)', () => {
    // This is the assertion that matters more than a bare rejection check:
    // confirm the payload isMintableTarget refuses is refused BECAUSE it
    // would have redeemed attachment A2 rather than the validated A, not for
    // some unrelated reason. Simulates performRequest's own concatenation
    // exactly (same prefix, same URL parser) rather than asserting against
    // the implementation's internals.
    const target = '/me/messages/M/attachments/A/%2e%2e/%2e%2e/attachments/A2/$value';
    const resolved = new URL(`https://graph.microsoft.com/v1.0${target}`);
    expect(resolved.pathname).toBe('/v1.0/me/messages/M/attachments/A2/$value');
    expect(resolved.pathname).not.toContain('/attachments/A/');
    expect(isMintableTarget(target)).toBe(false);
  });

  it('does not decode double-percent-encoded dot segments (%252e%252e stays literal, proven inert)', () => {
    // %252e is the percent-encoding of the literal string '%2e', not of '.'.
    // Nothing between this gate and fetch() ever decodes a percent-encoded
    // percent sign, so %252e%252e is NOT one of the URL Standard's six
    // dot-segment spellings and the parser does not resolve it away -- it
    // stays a literal (harmless, Graph-will-404-it) path segment rather than
    // collapsing into attachment A2. Confirmed by simulating performRequest's
    // own concatenation: the resolved pathname still contains the original
    // 'A' segment untouched, immediately before the literal encoded junk, so
    // no attachment substitution happens either way.
    const target = '/me/messages/M/attachments/A/%252e%252e/%252e%252e/attachments/A2/$value';
    const resolved = new URL(`https://graph.microsoft.com/v1.0${target}`);
    expect(resolved.pathname).toBe(
      '/v1.0/me/messages/M/attachments/A/%252e%252e/%252e%252e/attachments/A2/$value'
    );
    // Not rejected: it is inert against the mechanism this gate defends
    // against, and rejecting literal '%25' would cost real targets for no
    // safety benefit (percent signs are valid, if unusual, in Graph paths).
    expect(isMintableTarget(target)).toBe(true);
  });

  it('still mints legitimate targets containing literal (non-dot-segment) percent sequences', () => {
    // A filename that happens to already be percent-encoded by the caller,
    // e.g. copied verbatim from a @microsoft.graph.downloadUrl-adjacent
    // metadata field, must not be rejected just for containing '%'.
    expect(isMintableTarget('/me/drive/root:/Invoice%20Q3.pdf:/content')).toBe(true);
  });
});

/**
 * End-to-end proof for the exact Codex scenario: a redeemed ticket must not
 * reach the mail attachment's unsuffixed metadata endpoint. Exercised through
 * read-document, which is the tool named in the finding.
 */
describe('read-document refuses to mint for a target that diverges from its own pathname', () => {
  const URL_CONFIG = {
    base: 'http://m365-max-mcp:3001',
    key: 'shared-hmac-key',
    keyId: 'k1',
    ttlSeconds: 120,
  };
  let store: AttachmentTicketStore;

  function stubProxy(reply: ConvertResult): {
    client: AttachmentProxyClient;
    requests: Array<{ uri: string }>;
  } {
    const requests: Array<{ uri: string }> = [];
    const client = {
      convertToMarkdown: async (req: { uri: string }) => {
        requests.push(req);
        return reply;
      },
    } as unknown as AttachmentProxyClient;
    return { client, requests };
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

  it('refuses the exact PoC target as invalid_target, minting nothing and dialling nothing', async () => {
    const readDocument = UTILITY_TOOLS.find((t) => t.name === 'read-document')!;
    const { client: proxy, requests } = stubProxy({ ok: true, markdown: 'never reached' });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });

    const result = await readDocument.execute(
      { target: '/me/messages/M/attachments/A#/$value' },
      {
        graphClient: { makeRequest: vi.fn() } as never,
        authManager: {
          isOAuthModeEnabled: () => false,
          isMultiAccount: async () => false,
          getTokenForAccount: async () => 'SERVER_OWN_TOKEN',
        } as never,
        multiAccount: false,
        accountNames: [],
      }
    );

    expect(result.isError).toBe(true);
    const body = JSON.parse((result.content[0] as { text: string }).text);
    expect(body.error).toBe('invalid_target');
    expect(store.size()).toBe(0);
    expect(requests).toHaveLength(0);
  });

  it('refuses the Codex P1 #2 percent-encoded dot-segment payload, minting nothing and dialling nothing', async () => {
    const readDocument = UTILITY_TOOLS.find((t) => t.name === 'read-document')!;
    const { client: proxy, requests } = stubProxy({ ok: true, markdown: 'never reached' });
    configureAttachmentProxy({ client: proxy, url: 'http://docglean:8080/mcp' });

    const result = await readDocument.execute(
      { target: '/me/messages/M/attachments/A/%2e%2e/%2e%2e/attachments/A2/$value' },
      {
        graphClient: { makeRequest: vi.fn() } as never,
        authManager: {
          isOAuthModeEnabled: () => false,
          isMultiAccount: async () => false,
          getTokenForAccount: async () => 'SERVER_OWN_TOKEN',
        } as never,
        multiAccount: false,
        accountNames: [],
      }
    );

    expect(result.isError).toBe(true);
    const body = JSON.parse((result.content[0] as { text: string }).text);
    expect(body.error).toBe('invalid_target');
    // The real damage this closes: no ticket for A2 (or anything else) is
    // minted just because the string named A -- nothing is minted at all.
    expect(store.size()).toBe(0);
    expect(requests).toHaveLength(0);
  });
});

/**
 * get-download-url is a SECOND, independent call site that decides whether a
 * target may be minted (it does not go through `isMintableTarget` at all --
 * it re-implements the same family-match-plus-suffix check inline). Two
 * gates that can disagree is exactly the shape this fix must not leave
 * behind, so the same evasions must be proven closed here too.
 */
describe('get-download-url refuses the same divergent targets', () => {
  const tool = UTILITY_TOOLS.find((t) => t.name === 'get-download-url')!;

  function ctx() {
    return {
      graphClient: {} as never,
      authManager: {
        isOAuthModeEnabled: () => false,
        isMultiAccount: async () => false,
        getTokenForAccount: async () => 'SERVER_OWN_TOKEN',
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

  it('refuses the exact PoC # target', async () => {
    const result = await tool.execute({ target: '/me/messages/M/attachments/A#/$value' }, ctx());
    expect(result.isError).toBe(true);
    expect(parse(result as never).downloadUrl).toBeUndefined();
  });

  it('refuses a ? query-string target', async () => {
    const result = await tool.execute({ target: '/me/messages/M/attachments/A?x=/$value' }, ctx());
    expect(result.isError).toBe(true);
    expect(parse(result as never).downloadUrl).toBeUndefined();
  });

  it('still mints a legitimate mail attachment target', async () => {
    const result = await tool.execute({ target: LEGIT_MAIL_ATTACHMENT }, ctx());
    expect(parse(result as never).downloadUrl).toMatch(/^http:\/\/m365:3000\/attachment\?/);
  });

  it('refuses the Codex P1 #2 percent-encoded dot-segment payload', async () => {
    const result = await tool.execute(
      { target: '/me/messages/M/attachments/A/%2e%2e/%2e%2e/attachments/A2/$value' },
      ctx()
    );
    expect(result.isError).toBe(true);
    expect(parse(result as never).downloadUrl).toBeUndefined();
  });
});
