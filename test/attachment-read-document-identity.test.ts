import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { UTILITY_TOOLS } from '../src/graph-tools.js';
import { requestContext } from '../src/request-context.js';
import { AttachmentTicketStore } from '../src/lib/attachment-tickets.js';
import {
  configureAttachmentMinting,
  resetAttachmentMinting,
} from '../src/lib/attachment-minting.js';
import {
  configureAttachmentProxy,
  resetAttachmentProxy,
} from '../src/lib/attachment-proxy-runtime.js';

/**
 * read-document mints too, but unlike get-download-url it does not follow a request token into
 * the ticket: it refuses, so a document is never fetched as a different identity than the one
 * that asked for it. Both halves of the predicate are pinned, because `isOAuthModeEnabled()` is
 * false in plain bearer mode and in --obo, both of which still run inside a request context
 * holding the caller's token.
 */
describe('read-document refuses whenever Graph identity comes from the request', () => {
  const tool = UTILITY_TOOLS.find((t) => t.name === 'read-document')!;
  const MAIL_ATTACHMENT = '/me/messages/AAA/attachments/BBB/$value';
  let store: AttachmentTicketStore;
  let convert: ReturnType<typeof vi.fn>;

  function ctx(isOAuth = false) {
    return {
      graphClient: { makeRequest: vi.fn(async () => null) } as never,
      authManager: {
        isOAuthModeEnabled: () => isOAuth,
        isMultiAccount: async () => false,
        getToken: async () => 'SERVER_OWN_TOKEN',
        getTokenForAccount: async () => 'SERVER_OWN_TOKEN',
      } as never,
      multiAccount: false,
      accountNames: [],
    };
  }

  function errorOf(result: { content: Array<{ text: string }> }) {
    return JSON.parse(result.content[0].text).error;
  }

  beforeEach(() => {
    store = new AttachmentTicketStore(120);
    configureAttachmentMinting({
      store,
      config: { base: 'http://m365:3000', key: 'k', keyId: '1', ttlSeconds: 120 },
    });
    convert = vi.fn(async () => ({ ok: true, markdown: 'text' }));
    configureAttachmentProxy({
      client: { convertToMarkdown: convert } as never,
      url: 'http://docglean:8080/mcp',
    });
  });

  afterEach(() => {
    resetAttachmentMinting();
    resetAttachmentProxy();
  });

  it('reads when identity is the server’s own cached token', async () => {
    const result = await tool.execute({ target: MAIL_ATTACHMENT }, ctx());
    expect(result.isError).toBeFalsy();
    expect(convert).toHaveBeenCalledTimes(1);
  });

  it('refuses in bearer/OBO mode, where a request token is present, minting nothing', async () => {
    const result = await requestContext.run({ accessToken: 'CALLER_TOKEN' }, () =>
      tool.execute({ target: MAIL_ATTACHMENT }, ctx())
    );
    expect(result.isError).toBe(true);
    expect(errorOf(result as never)).toBe('identity_not_supported');
    expect(store.size()).toBe(0);
    expect(convert).not.toHaveBeenCalled();
  });

  it('refuses when OAuth mode is on, minting nothing', async () => {
    const result = await tool.execute({ target: MAIL_ATTACHMENT }, ctx(true));
    expect(result.isError).toBe(true);
    expect(errorOf(result as never)).toBe('identity_not_supported');
    expect(store.size()).toBe(0);
    expect(convert).not.toHaveBeenCalled();
  });
});
