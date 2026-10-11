import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mintDownloadUrl, UTILITY_TOOLS } from '../src/graph-tools.js';
import { requestContext } from '../src/request-context.js';
import { AttachmentTicketStore } from '../src/lib/attachment-tickets.js';
import {
  configureAttachmentMinting,
  resetAttachmentMinting,
} from '../src/lib/attachment-minting.js';

/**
 * The content-type probe is opt-in at mint time and only ever asks Graph about a single mail
 * or event attachment's /$value, after the target has passed the plain-path check, so a
 * target whose validated string and fetched path diverge is refused before Graph is asked
 * anything about it.
 */
describe('mintDownloadUrl content-type probe', () => {
  const MAIL = '/me/messages/AAA/attachments/BBB/$value';
  let store: AttachmentTicketStore;

  beforeEach(() => {
    store = new AttachmentTicketStore(120);
    configureAttachmentMinting({
      store,
      config: { base: 'http://127.0.0.1:1', key: Buffer.from('k'), keyId: 'k', ttlSeconds: 120 },
    } as never);
  });
  afterEach(() => resetAttachmentMinting());

  const authManager = {
    isOAuthModeEnabled: () => false,
    isMultiAccount: async () => false,
    getToken: async () => 'SERVER_TOKEN',
    getTokenForAccount: async () => 'SERVER_TOKEN',
  } as never;

  function graph(meta: unknown = { name: 'x.eml', contentType: 'message/rfc822', size: 9 }) {
    return { makeRequest: vi.fn(async () => meta) };
  }

  /** The ticket the mint produced, read back out of the store. */
  function mintedTicket() {
    const internal = store as unknown as { tickets: Map<string, unknown> };
    return store.redeem([...internal.tickets.keys()][0]);
  }

  it('does not ask Graph anything unless the caller opts in', async () => {
    const result = await mintDownloadUrl(MAIL, undefined, authManager);
    expect(result?.isError).toBeUndefined();
    expect(mintedTicket()).not.toHaveProperty('probedContentType');
  });

  it('get-download-url itself never probes', async () => {
    const tool = UTILITY_TOOLS.find((t) => t.name === 'get-download-url')!;
    const g = graph();
    const result = await tool.execute({ target: MAIL }, {
      graphClient: g as never,
      authManager,
      multiAccount: false,
    } as never);
    expect(result.isError).toBeUndefined();
    expect(g.makeRequest).not.toHaveBeenCalled();
  });

  it('records what Graph says about the attachment on the ticket', async () => {
    const g = graph();
    await mintDownloadUrl(MAIL, undefined, authManager, { probeWith: g });
    expect(g.makeRequest).toHaveBeenCalledWith(
      '/me/messages/AAA/attachments/BBB?$select=name,contentType',
      { accessToken: 'SERVER_TOKEN' }
    );
    expect(mintedTicket()).toMatchObject({
      probedContentType: 'message/rfc822',
      probedName: 'x.eml',
    });
  });

  it('probes an event attachment too, and with the request token when identity is the request', async () => {
    const g = graph();
    await requestContext.run({ accessToken: 'CALLER_TOKEN' } as never, () =>
      mintDownloadUrl('/me/events/E/attachments/A/$value', undefined, authManager, {
        probeWith: g,
      })
    );
    expect(g.makeRequest).toHaveBeenCalledWith(
      '/me/events/E/attachments/A?$select=name,contentType',
      {
        accessToken: 'CALLER_TOKEN',
      }
    );
  });

  it.each([
    '/me/photo/$value',
    '/me/drive/items/I/content',
    '/me/onlineMeetings/M/recordings/R/content',
    '/me/messages/M/attachments/A/B/$value',
  ])('never probes a target that is not a mail or event attachment: %s', async (target) => {
    const g = graph();
    const result = await mintDownloadUrl(target, undefined, authManager, { probeWith: g });
    expect(result?.isError).toBeUndefined();
    expect(g.makeRequest).not.toHaveBeenCalled();
  });

  it.each([
    '/me/messages/M/attachments/A#/$value',
    '/me/messages/M/attachments/A?x=/$value',
    '/me/messages/M/attachments/A/../../attachments/A2/$value',
    '/me/messages/M/attachments/A\t/$value',
  ])('refuses a divergent target before probing it: %s', async (target) => {
    const g = graph();
    const result = await mintDownloadUrl(target, undefined, authManager, { probeWith: g });
    expect(result?.isError).toBe(true);
    expect(g.makeRequest).not.toHaveBeenCalled();
    expect(store.size()).toBe(0);
  });

  it('mints without a hint when the probe fails or Graph has nothing specific', async () => {
    const failing = { makeRequest: vi.fn(async () => Promise.reject(new Error('403'))) };
    expect(
      (await mintDownloadUrl(MAIL, undefined, authManager, { probeWith: failing }))?.isError
    ).toBeUndefined();
    expect(mintedTicket()).not.toHaveProperty('probedContentType');

    const nulls = graph({ name: null, contentType: null });
    await mintDownloadUrl(MAIL, undefined, authManager, { probeWith: nulls });
    const ticket = mintedTicket();
    expect(ticket).not.toHaveProperty('probedContentType');
    expect(ticket).not.toHaveProperty('probedName');
  });

  it('composes with a multi-fetch count', async () => {
    await mintDownloadUrl(MAIL, undefined, authManager, { probeWith: graph(), maxRedemptions: 2 });
    const internal = store as unknown as { tickets: Map<string, unknown> };
    const id = [...internal.tickets.keys()][0];
    store.redeem(id)!.release();
    expect(store.redeem(id)).toMatchObject({ probedContentType: 'message/rfc822' });
  });
});
