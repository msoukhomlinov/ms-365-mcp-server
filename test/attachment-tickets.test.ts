import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { Writable } from 'node:stream';
import { UTILITY_TOOLS } from '../src/graph-tools.js';
import {
  configureAttachmentMinting,
  resetAttachmentMinting,
} from '../src/lib/attachment-minting.js';
import {
  AttachmentTicketStore,
  MAX_REDEMPTIONS,
  TicketStoreFullError,
  buildAttachmentUrl,
  TICKET_PARAM,
} from '../src/lib/attachment-tickets.js';
import {
  loadAttachmentUrlConfig,
  AttachmentUrlConfigError,
} from '../src/lib/attachment-url-config.js';
import { canonicalString, digest } from '../src/lib/url-signing.js';
import { createAttachmentHandler } from '../src/attachment-route.js';

describe('AttachmentTicketStore', () => {
  const NOW = 1_780_000_000_000;

  it('mints an id that is not guessable from the target', () => {
    const store = new AttachmentTicketStore(120);
    const a = store.mint('/me/messages/1/attachments/2/$value', undefined, NOW);
    const b = store.mint('/me/messages/1/attachments/2/$value', undefined, NOW);
    expect(a.id).not.toBe(b.id);
    // 32 bytes of base64url, unpadded.
    expect(a.id).toMatch(/^[A-Za-z0-9_-]{43}$/);
  });

  it('redeems exactly MAX_REDEMPTIONS times and then no more', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint('/target', 'max@example.com', NOW);
    for (let i = 0; i < MAX_REDEMPTIONS; i += 1) {
      const lease = store.redeem(id, NOW);
      expect(lease?.target, `redemption ${i + 1} of ${MAX_REDEMPTIONS}`).toBe('/target');
      expect(lease?.remaining).toBe(MAX_REDEMPTIONS - i - 1);
      lease!.release();
    }
    expect(store.redeem(id, NOW)).toBeUndefined();
  });

  it('drops the entry as the last redemption is handed out, not after', () => {
    // "Present in the map" and "redeemable" have to stay the same fact: an
    // entry left at zero would be a live-looking ticket that always refuses,
    // and it would hold a MAX_LIVE_TICKETS slot for the rest of its TTL.
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint('/target', undefined, NOW);
    for (let i = 0; i < MAX_REDEMPTIONS - 1; i += 1) store.redeem(id, NOW)!.release();
    expect(store.size(NOW)).toBe(1);
    store.redeem(id, NOW)!.release();
    expect(store.size(NOW)).toBe(0);
  });

  it('carries the account through to every redemption, not just the first', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint('/target', 'olga@example.com', NOW);
    const first = store.redeem(id, NOW)!;
    expect(first.accountName).toBe('olga@example.com');
    first.release();
    expect(store.redeem(id, NOW)?.accountName).toBe('olga@example.com');
  });

  it('refuses a ticket that has expired', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint('/target', undefined, NOW);
    expect(store.redeem(id, NOW + 119_000)).toBeDefined();

    const { id: second } = store.mint('/target', undefined, NOW);
    expect(store.redeem(second, NOW + 121_000)).toBeUndefined();
  });

  it('treats the expiry instant itself as dead', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint('/target', undefined, NOW);
    expect(store.redeem(id, NOW + 120_000)).toBeUndefined();
  });

  it('lets expiry beat an unspent budget', () => {
    // The budget shortens a ticket's life; it must never extend it. A ticket
    // with every redemption still in hand is dead the instant its TTL passes.
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint('/target', undefined, NOW);
    expect(store.redemptionsLeft(id, NOW)).toBe(MAX_REDEMPTIONS);
    expect(store.redeem(id, NOW + 120_000)).toBeUndefined();
    expect(store.redemptionsLeft(id, NOW)).toBe(0);
  });

  it('expires a ticket mid-budget rather than letting it run out its fetches', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint('/target', undefined, NOW);
    store.redeem(id, NOW)!.release();
    expect(store.redemptionsLeft(id, NOW)).toBe(MAX_REDEMPTIONS - 1);
    expect(store.redeem(id, NOW + 121_000)).toBeUndefined();
  });

  it('gives the same undefined for unknown, exhausted and expired', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint('/target', undefined, NOW);
    for (let i = 0; i < MAX_REDEMPTIONS; i += 1) store.redeem(id, NOW)!.release();
    expect(store.redeem(id, NOW)).toBeUndefined();
    expect(store.redeem('never-existed', NOW)).toBeUndefined();
    expect(store.redeem('', NOW)).toBeUndefined();

    const { id: unexpired } = store.mint('/target', undefined, NOW);
    expect(store.redeem(unexpired, NOW + 121_000)).toBeUndefined();
  });

  it('refuses a ticket another redemption is holding, indistinguishably', () => {
    // The in-flight refusal is the only one an attacker can provoke on demand,
    // by racing a legitimate fetch. It has to be the same `undefined` as an id
    // that never existed, or racing becomes a way to confirm a guess.
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint('/target', undefined, NOW);
    const held = store.redeem(id, NOW)!;
    expect(store.redeem(id, NOW)).toBe(store.redeem('never-existed', NOW));
    expect(store.redeem(id, NOW)).toBeUndefined();
    // ...and the hold is temporary, not a burn: releasing restores it.
    held.release();
    expect(store.redeem(id, NOW)).toBeDefined();
  });

  it('does not spend a redemption on a request it refused for being in flight', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint('/target', undefined, NOW);
    const held = store.redeem(id, NOW)!;
    for (let i = 0; i < 10; i += 1) expect(store.redeem(id, NOW)).toBeUndefined();
    held.release();
    expect(store.redemptionsLeft(id, NOW)).toBe(MAX_REDEMPTIONS - 1);
  });

  it('treats release as idempotent', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint('/target', undefined, NOW);
    const lease = store.redeem(id, NOW)!;
    lease.release();
    const second = store.redeem(id, NOW)!;
    // A stale release from the first lease must not unlock the second one.
    lease.release();
    expect(store.redeem(id, NOW)).toBeUndefined();
    second.release();
    expect(store.redeem(id, NOW)).toBeDefined();
  });

  it('survives releasing a lease whose ticket is already gone', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint('/target', undefined, NOW);
    let last!: ReturnType<AttachmentTicketStore['redeem']>;
    for (let i = 0; i < MAX_REDEMPTIONS; i += 1) {
      last = store.redeem(id, NOW);
      if (i < MAX_REDEMPTIONS - 1) last!.release();
    }
    // The final redemption removed the entry while still holding this lease.
    expect(store.size(NOW)).toBe(0);
    expect(() => last!.release()).not.toThrow();
    expect(store.redeem(id, NOW)).toBeUndefined();
  });

  it('lets an expiry sweep collect a ticket a lease is still holding', () => {
    // A stream can outlive the TTL. The sweep must not be blocked by the hold,
    // and releasing afterwards must not resurrect anything.
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint('/target', undefined, NOW);
    const held = store.redeem(id, NOW)!;
    expect(store.size(NOW + 121_000)).toBe(0);
    held.release();
    expect(store.redeem(id, NOW + 121_000)).toBeUndefined();
    expect(store.size(NOW + 121_000)).toBe(0);
  });

  it('sweeps expired tickets rather than letting them accumulate', () => {
    const store = new AttachmentTicketStore(60);
    for (let i = 0; i < 10; i += 1) store.mint(`/t${i}`, undefined, NOW);
    expect(store.size(NOW)).toBe(10);
    expect(store.size(NOW + 61_000)).toBe(0);
  });

  it('refuses to mint past the live-ticket cap instead of growing without bound', () => {
    const store = new AttachmentTicketStore(300);
    for (let i = 0; i < 256; i += 1) store.mint(`/t${i}`, undefined, NOW);
    expect(() => store.mint('/one-too-many', undefined, NOW)).toThrow(TicketStoreFullError);
  });

  it('frees slots once the outstanding tickets expire', () => {
    const store = new AttachmentTicketStore(60);
    for (let i = 0; i < 256; i += 1) store.mint(`/t${i}`, undefined, NOW);
    expect(() => store.mint('/blocked', undefined, NOW)).toThrow(TicketStoreFullError);
    expect(() => store.mint('/ok', undefined, NOW + 61_000)).not.toThrow();
  });

  it('never evicts a live ticket to make room', () => {
    const store = new AttachmentTicketStore(300);
    const first = store.mint('/keep-me', undefined, NOW);
    for (let i = 0; i < 255; i += 1) store.mint(`/t${i}`, undefined, NOW);
    expect(() => store.mint('/overflow', undefined, NOW)).toThrow(TicketStoreFullError);
    // The oldest ticket is still redeemable — a caller minting in a loop must
    // not be able to invalidate someone else's outstanding ticket.
    expect(store.redeem(first.id, NOW)).toBeDefined();
  });

  it('still refuses at the cap when tickets are partly redeemed, not evicting', () => {
    // Regression guard for the budget: a ticket with redemptions left is live
    // and holds its slot, so the cap must be reached and refused the same way.
    // Evicting a partly-used ticket would be worse than evicting a fresh one —
    // the holder has already started a multi-fetch flow against it.
    const store = new AttachmentTicketStore(300);
    const first = store.mint('/keep-me', undefined, NOW);
    store.redeem(first.id, NOW)!.release();
    for (let i = 0; i < 255; i += 1) {
      const t = store.mint(`/t${i}`, undefined, NOW);
      store.redeem(t.id, NOW)!.release();
    }
    expect(store.size(NOW)).toBe(256);
    expect(() => store.mint('/overflow', undefined, NOW)).toThrow(TicketStoreFullError);
    expect(store.redemptionsLeft(first.id, NOW)).toBe(MAX_REDEMPTIONS - 1);
  });

  it('frees a slot as soon as a ticket is exhausted, without waiting for the TTL', () => {
    const store = new AttachmentTicketStore(300);
    const ids: string[] = [];
    for (let i = 0; i < 256; i += 1) ids.push(store.mint(`/t${i}`, undefined, NOW).id);
    expect(() => store.mint('/blocked', undefined, NOW)).toThrow(TicketStoreFullError);

    for (let i = 0; i < MAX_REDEMPTIONS; i += 1) store.redeem(ids[0], NOW)!.release();
    expect(store.size(NOW)).toBe(255);
    expect(() => store.mint('/now-fits', undefined, NOW)).not.toThrow();
  });
});

describe('buildAttachmentUrl', () => {
  const config = { base: 'http://m365-max-mcp:3000', key: 'k', keyId: '1', ttlSeconds: 120 };

  it('puts the ticket in the query, never the path', () => {
    const url = new URL(buildAttachmentUrl(config, 'TICKET', 1_780_000_000_000));
    expect(url.pathname).toBe('/attachment');
    expect(url.pathname).not.toContain('TICKET');
    expect(url.searchParams.get(TICKET_PARAM)).toBe('TICKET');
  });

  it('mints a signature that verifies over its own canonical string', () => {
    const url = buildAttachmentUrl(config, 'TICKET', 1_780_000_000_000);
    const parsed = new URL(url);
    expect(digest(config.key, canonicalString(url, parsed.searchParams.get('dgx')!))).toBe(
      parsed.searchParams.get('dgs')
    );
  });

  it('honours the configured key id', () => {
    const url = new URL(
      buildAttachmentUrl({ ...config, keyId: 'rotated' }, 'T', 1_780_000_000_000)
    );
    expect(url.searchParams.get('dgk')).toBe('rotated');
  });
});

describe('loadAttachmentUrlConfig', () => {
  const good = {
    MS365_MCP_ATTACHMENT_URL_BASE: 'http://m365-max-mcp:3000',
    MS365_MCP_ATTACHMENT_URL_KEY: 'secret',
  } as Record<string, string | undefined>;

  it('returns null when the feature is off, whatever the environment says', () => {
    expect(loadAttachmentUrlConfig(false, good)).toBeNull();
  });

  it('applies the documented defaults', () => {
    const config = loadAttachmentUrlConfig(true, good)!;
    expect(config.keyId).toBe('1');
    expect(config.ttlSeconds).toBe(120);
    expect(config.base).toBe('http://m365-max-mcp:3000');
  });

  it('refuses to start without a base', () => {
    expect(() => loadAttachmentUrlConfig(true, { MS365_MCP_ATTACHMENT_URL_KEY: 'k' })).toThrow(
      AttachmentUrlConfigError
    );
  });

  it('refuses to start without a key', () => {
    expect(() =>
      loadAttachmentUrlConfig(true, { MS365_MCP_ATTACHMENT_URL_BASE: 'http://h:3000' })
    ).toThrow(AttachmentUrlConfigError);
  });

  it('refuses an IPv6-literal base, which could never sign compatibly', () => {
    // Regression: WHATWG URL.hostname keeps the brackets and rewrites some
    // literals into a compressed form; the verifier's Python does neither. A
    // base like this produced a healthy startup and a 100% refusal rate at the
    // far end, with nothing connecting the two.
    for (const base of ['http://[fd00::1]:3000', 'http://[::1]:3000']) {
      expect(() =>
        loadAttachmentUrlConfig(true, { ...good, MS365_MCP_ATTACHMENT_URL_BASE: base })
      ).toThrow(AttachmentUrlConfigError);
    }
  });

  it('caps the TTL at the verifier default rather than at a round number', () => {
    // Regression: the ceiling was 3600, but docglean refuses anything past its
    // own _MAX_TTL_S (300) plus _CLOCK_SKEW_S (5) — so every value from 306 to
    // 3600 was advertised as valid and was in fact unusable.
    expect(
      loadAttachmentUrlConfig(true, { ...good, MS365_MCP_ATTACHMENT_URL_TTL_S: '300' })!.ttlSeconds
    ).toBe(300);
    expect(() =>
      loadAttachmentUrlConfig(true, { ...good, MS365_MCP_ATTACHMENT_URL_TTL_S: '301' })
    ).toThrow(AttachmentUrlConfigError);
  });

  it('strips a key file the way Python does, not the way trim() does', async () => {
    // Regression: .trim() removes a BOM (Python does not) and leaves U+0085
    // (Python removes it, and it slips past the control-character guard because
    // it is not < 0x20). Either way both ends derive different key bytes.
    const { mkdtempSync, writeFileSync } = await import('node:fs');
    const { join } = await import('node:path');
    const { tmpdir } = await import('node:os');
    const dir = mkdtempSync(join(tmpdir(), 'dgkey-'));

    const bom = join(dir, 'bom');
    writeFileSync(bom, '﻿secret\n');
    expect(
      loadAttachmentUrlConfig(true, {
        MS365_MCP_ATTACHMENT_URL_BASE: 'http://h:3000',
        MS365_MCP_ATTACHMENT_URL_KEY_FILE: bom,
      })!.key
    ).toBe('﻿secret');

    const nel = join(dir, 'nel');
    writeFileSync(nel, 'secret');
    expect(
      loadAttachmentUrlConfig(true, {
        MS365_MCP_ATTACHMENT_URL_BASE: 'http://h:3000',
        MS365_MCP_ATTACHMENT_URL_KEY_FILE: nel,
      })!.key
    ).toBe('secret');
  });

  it('refuses a base carrying a query string', () => {
    expect(() =>
      loadAttachmentUrlConfig(true, {
        ...good,
        MS365_MCP_ATTACHMENT_URL_BASE: 'http://h:3000/?a=1',
      })
    ).toThrow(AttachmentUrlConfigError);
  });

  it('refuses a non-http scheme', () => {
    expect(() =>
      loadAttachmentUrlConfig(true, { ...good, MS365_MCP_ATTACHMENT_URL_BASE: 'ftp://h:3000' })
    ).toThrow(AttachmentUrlConfigError);
  });

  it('refuses a control character in the key without echoing it', () => {
    try {
      loadAttachmentUrlConfig(true, { ...good, MS365_MCP_ATTACHMENT_URL_KEY: 'ab\ncd' });
      throw new Error('should have thrown');
    } catch (error) {
      expect((error as Error).message).toContain('offset 2');
      expect((error as Error).message).not.toContain('ab');
    }
  });

  it('refuses a non-numeric TTL rather than silently defaulting', () => {
    expect(() =>
      loadAttachmentUrlConfig(true, { ...good, MS365_MCP_ATTACHMENT_URL_TTL_S: '12abc' })
    ).toThrow(AttachmentUrlConfigError);
  });

  it('refuses a TTL of zero or beyond the ceiling', () => {
    expect(() =>
      loadAttachmentUrlConfig(true, { ...good, MS365_MCP_ATTACHMENT_URL_TTL_S: '0' })
    ).toThrow(AttachmentUrlConfigError);
    expect(() =>
      loadAttachmentUrlConfig(true, { ...good, MS365_MCP_ATTACHMENT_URL_TTL_S: '3601' })
    ).toThrow(AttachmentUrlConfigError);
  });
});

describe('attachment redemption route', () => {
  let store: AttachmentTicketStore;
  let sent: { status?: number; body?: unknown; headers: Record<string, string> };
  let written: Buffer[];

  /**
   * A real `Writable`, not a bag of spies: the handler finishes by awaiting
   * `stream.pipeline(..., res)`, which resolves only on a genuine 'finish'.
   * A mock that merely records `write` calls leaves that promise pending and
   * the test times out rather than failing usefully.
   */
  function mockRes() {
    const res = new Writable({
      write(chunk: Buffer, _enc, cb) {
        written.push(Buffer.from(chunk));
        cb();
      },
    }) as Writable & Record<string, unknown>;
    res.status = (code: number) => {
      sent.status = code;
      return res;
    };
    res.type = () => res;
    res.send = (body: unknown) => {
      sent.body = body;
      return res;
    };
    res.setHeader = (name: string, value: string) => {
      sent.headers[name] = value;
    };
    return res;
  }

  const authManager = { isOAuthModeEnabled: () => false, getTokenForAccount: async () => 'tok' };

  beforeEach(() => {
    store = new AttachmentTicketStore(120);
    sent = { headers: {} };
    written = [];
  });

  it('404s a missing ticket parameter', async () => {
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () => ({}) as never,
      authManager: authManager as never,
    });
    await handler({ query: {} } as never, mockRes() as never, (() => {}) as never);
    expect(sent.status).toBe(404);
  });

  it('404s a repeated ticket parameter rather than picking one', async () => {
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () => ({}) as never,
      authManager: authManager as never,
    });
    const { id } = store.mint('/target', undefined);
    await handler(
      { query: { t: [id, 'guess'] } } as never,
      mockRes() as never,
      (() => {}) as never
    );
    expect(sent.status).toBe(404);
    // and the real ticket must survive an attempt that was refused
    expect(store.redeem(id)).toBeDefined();
  });

  it('404s unknown, exhausted, expired and in-flight tickets identically', async () => {
    /**
     * Drive the route once against its own store and return everything the
     * peer could observe. Anything that varies between these four is a probe
     * oracle: it tells a caller holding a guessed id that the id is real.
     */
    async function refusal(target: string, targetStore: AttachmentTicketStore) {
      const handler = createAttachmentHandler({
        store: targetStore,
        getGraphClient: () => ({}) as never,
        authManager: authManager as never,
      });
      sent = { headers: {} };
      written = [];
      await handler({ query: { t: target } } as never, mockRes() as never, (() => {}) as never);
      return { status: sent.status, body: sent.body, headers: { ...sent.headers }, bytes: written };
    }

    const unknown = await refusal('nope', store);
    expect(unknown.status).toBe(404);

    const exhausted = store.mint('/target', undefined).id;
    for (let i = 0; i < MAX_REDEMPTIONS; i += 1) store.redeem(exhausted)!.release();

    // A one-second TTL minted in the past, so a real clock expires it: the
    // route calls `redeem` with no `nowMs`, so an injected timestamp would
    // never reach it and the ticket would still be live.
    const shortStore = new AttachmentTicketStore(1);
    const expired = shortStore.mint('/target', undefined, Date.now() - 5_000).id;

    // Held, not spent: a second request arriving while a legitimate fetch of
    // the same ticket is still streaming.
    const inFlightStore = new AttachmentTicketStore(120);
    const held = inFlightStore.mint('/target', undefined).id;
    const hold = inFlightStore.redeem(held)!;

    for (const [label, target, targetStore] of [
      ['exhausted', exhausted, store],
      ['expired', expired, shortStore],
      ['in-flight', held, inFlightStore],
    ] as const) {
      expect(await refusal(target, targetStore), label).toEqual(unknown);
    }

    hold.release();
  });

  it('streams the Graph body, spends one redemption, and releases the hold', async () => {
    const downloadStream = vi.fn(async () => ({
      body: new ReadableStream({
        start(controller) {
          controller.enqueue(new Uint8Array([1, 2, 3]));
          controller.close();
        },
      }),
      contentType: 'application/pdf',
      contentLength: 3,
      contentDisposition: 'attachment; filename="q.pdf"',
    }));
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () => ({ downloadStream }) as never,
      authManager: authManager as never,
    });
    const { id } = store.mint('/me/messages/1/attachments/2/$value', 'max@example.com');

    await handler({ query: { t: id } } as never, mockRes() as never, (() => {}) as never);

    expect(downloadStream).toHaveBeenCalledWith('/me/messages/1/attachments/2/$value', {
      accessToken: 'tok',
    });
    expect(sent.status).toBe(200);
    expect(sent.headers['content-type']).toBe('application/pdf');
    expect(sent.headers['content-disposition']).toBe('attachment; filename="q.pdf"');
    expect(sent.headers['x-content-type-options']).toBe('nosniff');
    expect(sent.headers['cache-control']).toBe('no-store');
    expect(Buffer.concat(written)).toEqual(Buffer.from([1, 2, 3]));
    // One redemption gone, the rest still there, and — the part a missing
    // `finally` would break — the ticket is not left pinned as in-flight.
    expect(store.redemptionsLeft(id)).toBe(MAX_REDEMPTIONS - 1);
    expect(store.redeem(id)).toBeDefined();
  });

  it('forces a download disposition when Graph supplies none', async () => {
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () =>
        ({
          downloadStream: async () => ({
            body: new ReadableStream({
              start(c) {
                c.close();
              },
            }),
            contentType: 'text/html',
            contentLength: null,
            contentDisposition: null,
          }),
        }) as never,
      authManager: authManager as never,
    });
    const { id } = store.mint('/t', undefined);
    await handler({ query: { t: id } } as never, mockRes() as never, (() => {}) as never);
    expect(sent.headers['content-disposition']).toBe('attachment');
    expect(sent.headers['content-length']).toBeUndefined();
  });

  it('502s an upstream failure and leaves the URL fetchable again', async () => {
    // The live defect: a Graph 404 on /me/photo/$value answered 502 *and* left
    // the caller a dead URL, so the only recovery was a re-mint the tool never
    // mentioned. The redemption is still spent — but the ticket is not.
    let attempt = 0;
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () =>
        ({
          downloadStream: async () => {
            attempt += 1;
            if (attempt === 1) throw new Error('Microsoft Graph API error: 404');
            return {
              body: new ReadableStream({
                start(c) {
                  c.enqueue(new Uint8Array([7, 7]));
                  c.close();
                },
              }),
              contentType: 'image/jpeg',
              contentLength: 2,
              contentDisposition: null,
            };
          },
        }) as never,
      authManager: authManager as never,
    });
    const { id } = store.mint('/me/photo/$value', undefined);

    await handler({ query: { t: id } } as never, mockRes() as never, (() => {}) as never);
    expect(sent.status).toBe(502);
    expect(store.redemptionsLeft(id)).toBe(MAX_REDEMPTIONS - 1);

    // Retrying the *same* URL is what the tool now tells the agent to do.
    sent = { headers: {} };
    written = [];
    await handler({ query: { t: id } } as never, mockRes() as never, (() => {}) as never);
    expect(sent.status).toBe(200);
    expect(Buffer.concat(written)).toEqual(Buffer.from([7, 7]));
  });

  it('503s an uninitialised client and still leaves the URL fetchable', async () => {
    let client: unknown = null;
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () => client as never,
      authManager: authManager as never,
    });
    const { id } = store.mint('/t', undefined);

    await handler({ query: { t: id } } as never, mockRes() as never, (() => {}) as never);
    expect(sent.status).toBe(503);
    expect(store.redemptionsLeft(id)).toBe(MAX_REDEMPTIONS - 1);

    client = {
      downloadStream: async () => ({
        body: new ReadableStream({
          start(c) {
            c.enqueue(new Uint8Array([9]));
            c.close();
          },
        }),
        contentType: 'application/pdf',
        contentLength: 1,
        contentDisposition: null,
      }),
    };
    sent = { headers: {} };
    written = [];
    await handler({ query: { t: id } } as never, mockRes() as never, (() => {}) as never);
    expect(sent.status).toBe(200);
    expect(Buffer.concat(written)).toEqual(Buffer.from([9]));
  });

  it('spends a redemption on every fetch, so a failing target cannot loop forever', async () => {
    // The counterpart to the two tests above: failures are survivable, not
    // free. Refunding them would make `maxFetches` untrue and let a leaked URL
    // pointing at a permanently-broken target be re-fetched all TTL long.
    const downloadStream = vi.fn(async () => {
      throw new Error('Microsoft Graph API error: 404');
    });
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () => ({ downloadStream }) as never,
      authManager: authManager as never,
    });
    const { id } = store.mint('/gone', undefined);

    for (let i = 0; i < MAX_REDEMPTIONS; i += 1) {
      sent = { headers: {} };
      await handler({ query: { t: id } } as never, mockRes() as never, (() => {}) as never);
      expect(sent.status, `attempt ${i + 1}`).toBe(502);
    }

    sent = { headers: {} };
    await handler({ query: { t: id } } as never, mockRes() as never, (() => {}) as never);
    expect(sent.status).toBe(404);
    // Not merely "the response was a 404": the handler must never have reached
    // Graph a fourth time. A store that answered but a route that refused, or
    // the reverse, would both pass a status-only assertion.
    expect(downloadStream).toHaveBeenCalledTimes(MAX_REDEMPTIONS);
  });

  it('serves a probe and then a conversion from one minted URL', async () => {
    // docglean's own advice — probe a large document, then convert it — was
    // impossible: the probe consumed the ticket and the convert 404'd. Two
    // fetches of one URL, both served, is the defect's direct inverse.
    const downloadStream = vi.fn(async () => ({
      body: new ReadableStream({
        start(c) {
          c.enqueue(new Uint8Array([1, 2, 3, 4]));
          c.close();
        },
      }),
      contentType: 'application/pdf',
      contentLength: 4,
      contentDisposition: null,
    }));
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () => ({ downloadStream }) as never,
      authManager: authManager as never,
    });
    const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined);

    for (const phase of ['probe', 'convert']) {
      sent = { headers: {} };
      written = [];
      await handler({ query: { t: id } } as never, mockRes() as never, (() => {}) as never);
      expect(sent.status, phase).toBe(200);
      expect(Buffer.concat(written), phase).toEqual(Buffer.from([1, 2, 3, 4]));
    }
    expect(downloadStream).toHaveBeenCalledTimes(2);
    // And a continuation still has room, which is what the third one is for.
    expect(store.redemptionsLeft(id)).toBe(MAX_REDEMPTIONS - 2);
  });

  it('never streams one ticket twice at once, whichever request wins the race', async () => {
    // Node is single-threaded, but every `await` in the handler is a point
    // where the other request runs. `redeem` is synchronous precisely so that
    // taking the ticket and marking it in-flight cannot be split by one.
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    const downloadStream = vi.fn(async () => {
      // Suspend inside the first request, exactly where a real fetch would.
      await gate;
      return {
        body: new ReadableStream({
          start(c) {
            c.enqueue(new Uint8Array([5]));
            c.close();
          },
        }),
        contentType: 'application/pdf',
        contentLength: 1,
        contentDisposition: null,
      };
    });
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () => ({ downloadStream }) as never,
      authManager: authManager as never,
    });
    const { id } = store.mint('/t', undefined);

    // The shared `mockRes` writes into the suite-level `sent`/`written`, which
    // two overlapping requests would trample. These own theirs.
    function ownRes() {
      const seen: { status?: number; headers: Record<string, string> } = { headers: {} };
      const bytes: Buffer[] = [];
      const res = new Writable({
        write(chunk: Buffer, _enc, cb) {
          bytes.push(Buffer.from(chunk));
          cb();
        },
      }) as Writable & Record<string, unknown>;
      res.status = (code: number) => ((seen.status = code), res);
      res.type = () => res;
      res.send = () => res;
      res.setHeader = (name: string, value: string) => {
        seen.headers[name] = value;
      };
      return { res, seen, bytes };
    }

    const one = ownRes();
    const first = handler({ query: { t: id } } as never, one.res as never, (() => {}) as never);

    // The second arrives while the first is parked inside downloadStream.
    await Promise.resolve();
    const two = ownRes();
    await handler({ query: { t: id } } as never, two.res as never, (() => {}) as never);

    // Refused, and refused before Graph was asked a second time — the strong
    // form: a 404 alone would also be produced by a route that called Graph
    // and then threw the bytes away.
    expect(two.seen.status).toBe(404);
    expect(two.bytes).toEqual([]);
    expect(downloadStream).toHaveBeenCalledTimes(1);

    release();
    await first;
    expect(one.seen.status).toBe(200);
    expect(Buffer.concat(one.bytes)).toEqual(Buffer.from([5]));

    // The loser is refused, not punished: the hold cost no redemption, and the
    // ticket is usable again now the winner has finished.
    expect(store.redemptionsLeft(id)).toBe(MAX_REDEMPTIONS - 1);
  });

  it('releases the hold even when the stream aborts mid-body', async () => {
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () =>
        ({
          downloadStream: async () => ({
            body: new ReadableStream({
              start(c) {
                c.enqueue(new Uint8Array([1]));
                c.error(new Error('connection reset'));
              },
            }),
            contentType: 'application/pdf',
            contentLength: null,
            contentDisposition: null,
          }),
        }) as never,
      authManager: authManager as never,
    });
    const { id } = store.mint('/t', undefined);
    await handler({ query: { t: id } } as never, mockRes() as never, (() => {}) as never);

    expect(sent.status).toBe(200);
    // An abort spent its redemption — bytes did leave this server, so it is
    // not a failure that delivered nothing — but it must not pin the ticket
    // as in-flight for the rest of its TTL.
    expect(store.redemptionsLeft(id)).toBe(MAX_REDEMPTIONS - 1);
    expect(store.redeem(id)).toBeDefined();
  });
});

/**
 * The tool output is the interface the model actually reads.
 *
 * Both defects this suite covers were, from an agent's point of view, defects
 * in this text: a URL described as "valid for one fetch" is one an agent will
 * never retry and never probe first, and a converter's own `fetch_failed`
 * carries nothing to say whether the URL is retryable, finished, or was never
 * valid. Changing the lifecycle without changing what the tool says would fix
 * nothing the agents can observe.
 */
describe('get-download-url states the lifecycle it actually has', () => {
  const tool = UTILITY_TOOLS.find((t) => t.name === 'get-download-url')!;
  const MAIL_ATTACHMENT = '/me/messages/AAA/attachments/BBB/$value';

  const ctx = {
    graphClient: {} as never,
    authManager: {
      isOAuthModeEnabled: () => false,
      isMultiAccount: async () => false,
      getTokenForAccount: async () => 'SERVER_OWN_TOKEN',
    } as never,
    multiAccount: false,
    accountNames: [],
  };

  beforeEach(() => {
    configureAttachmentMinting({
      store: new AttachmentTicketStore(120),
      config: { base: 'http://m365:3000', key: 'k', keyId: '1', ttlSeconds: 120 },
    });
  });

  afterEach(() => resetAttachmentMinting());

  async function mint() {
    const result = await tool.execute({ target: MAIL_ATTACHMENT }, ctx);
    return JSON.parse((result as { content: Array<{ text: string }> }).content[0].text);
  }

  it('advertises the real fetch budget rather than single use', async () => {
    const body = await mint();
    expect(body.maxFetches).toBe(MAX_REDEMPTIONS);
    // Stated, not omitted: the field used to say `true`, and an absent field
    // reads as "unknown, assume the old rule".
    expect(body.singleUse).toBe(false);
    expect(body.note).toContain(String(MAX_REDEMPTIONS));
    expect(body.note).not.toMatch(/valid for one fetch/i);
  });

  it('keeps the advertised count and the enforced count the same number', async () => {
    // The two used to be independent strings. Deriving both from
    // MAX_REDEMPTIONS is only worth anything if nothing re-states it.
    const body = await mint();
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint(MAIL_ATTACHMENT, undefined);
    let served = 0;
    for (let attempt = 0; attempt < 20; attempt += 1) {
      const lease = store.redeem(id);
      if (!lease) break;
      served += 1;
      lease.release();
    }
    expect(served).toBe(body.maxFetches);
  });

  it('says a failed fetch is retried on the same URL, not re-minted', async () => {
    const body = await mint();
    expect(body.onFetchFailure).toBeTypeOf('string');
    expect(body.onFetchFailure).toMatch(/same downloadUrl/i);
    expect(body.onFetchFailure).toMatch(/404/);
    expect(body.note).toMatch(/fails|failed/i);
  });

  it('still says what it always did about provenance and expiry', async () => {
    const body = await mint();
    expect(body.note).toContain('not by Microsoft Graph');
    expect(body.expiresAt).toMatch(/^\d{4}-\d{2}-\d{2}T/);
    expect(body.note).toContain('expiresAt');
  });
});
