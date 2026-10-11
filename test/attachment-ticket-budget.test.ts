import { describe, it, expect, vi } from 'vitest';
import { Writable } from 'node:stream';
import {
  AttachmentTicketStore,
  MAX_REDEMPTIONS,
  TicketStoreFullError,
} from '../src/lib/attachment-tickets.js';
import { createAttachmentHandler, createAttachmentUploadHandler } from '../src/attachment-route.js';
import { mintDownloadUrl } from '../src/graph-tools.js';
import {
  configureAttachmentMinting,
  resetAttachmentMinting,
} from '../src/lib/attachment-minting.js';

const NOW = 1_780_000_000_000;
const TARGET = '/me/messages/1/attachments/2/$value';

describe('store: per-ticket redemption count', () => {
  it('defaults to a single use, burnt as it is handed out', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint(TARGET, undefined, NOW);
    expect(store.redeem(id, NOW)).toBeDefined();
    expect(store.size(NOW)).toBe(0);
    expect(store.redeem(id, NOW)).toBeUndefined();
  });

  it('serves exactly maxRedemptions fetches and then refuses, per ticket', () => {
    const store = new AttachmentTicketStore(120);
    const multi = store.mint(TARGET, undefined, NOW, MAX_REDEMPTIONS);
    const single = store.mint(TARGET, undefined, NOW);
    for (let i = 0; i < MAX_REDEMPTIONS; i += 1) {
      store.redeem(multi.id, NOW)!.release();
    }
    expect(store.redeem(multi.id, NOW)).toBeUndefined();
    // The budget belongs to the ticket: the single-use one is untouched by the multi-use one.
    expect(store.redeem(single.id, NOW)).toBeDefined();
    expect(store.redeem(single.id, NOW)).toBeUndefined();
  });

  it('gives each ticket its own full budget, not a shared one', () => {
    const store = new AttachmentTicketStore(120);
    const a = store.mint(TARGET, undefined, NOW, 3);
    const b = store.mint(TARGET, undefined, NOW, 3);
    store.redeem(a.id, NOW)!.release();
    store.redeem(a.id, NOW)!.release();
    for (let i = 0; i < 3; i += 1) store.redeem(b.id, NOW)!.release();
    expect(store.redeem(b.id, NOW)).toBeUndefined();
    expect(store.redeem(a.id, NOW)).toBeDefined();
  });

  it('drops the entry as the last redemption is handed out, freeing the slot', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint(TARGET, undefined, NOW, 2);
    store.redeem(id, NOW)!.release();
    expect(store.size(NOW)).toBe(1);
    store.redeem(id, NOW);
    expect(store.size(NOW)).toBe(0);
  });

  it('refuses a ticket another redemption holds, and the refusal costs nothing', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint(TARGET, undefined, NOW, 2);
    const lease = store.redeem(id, NOW)!;
    expect(store.redeem(id, NOW)).toBeUndefined();
    expect(store.redeem(id, NOW)).toBeUndefined();
    lease.release();
    expect(store.redeem(id, NOW)).toBeDefined();
    expect(store.redeem(id, NOW)).toBeUndefined();
  });

  it('treats release as idempotent and never drops a later lease', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint(TARGET, undefined, NOW, 3);
    const first = store.redeem(id, NOW)!;
    first.release();
    const second = store.redeem(id, NOW)!;
    first.release();
    expect(store.redeem(id, NOW)).toBeUndefined();
    second.release();
  });

  it('lets the TTL end a ticket with budget left', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mint(TARGET, undefined, NOW, 3);
    store.redeem(id, NOW + 60_000)!.release();
    expect(store.redeem(id, NOW + 120_000)).toBeUndefined();
  });

  it('keeps the capacity refusal, counting a part-used ticket as live', () => {
    const store = new AttachmentTicketStore(120);
    for (let i = 0; i < 256; i += 1) store.mint(`/t${i}`, undefined, NOW, 3);
    expect(() => store.mint('/one-too-many', undefined, NOW)).toThrow(TicketStoreFullError);
  });

  it('keeps identity and purpose through every redemption', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mintWithToken(TARGET, 'CALLER_TOKEN', NOW, 2);
    for (let i = 0; i < 2; i += 1) {
      const lease = store.redeem(id, NOW)!;
      expect(lease).toMatchObject({
        kind: 'request-token',
        accessToken: 'CALLER_TOKEN',
        purpose: 'download',
      });
      expect(lease).not.toHaveProperty('accountName');
      lease.release();
    }
  });

  it('rejects a count outside 1..MAX_REDEMPTIONS', () => {
    const store = new AttachmentTicketStore(120);
    for (const bad of [0, -1, 1.5, MAX_REDEMPTIONS + 1, NaN]) {
      expect(() => store.mint(TARGET, undefined, NOW, bad)).toThrow(RangeError);
    }
    expect(store.size(NOW)).toBe(0);
  });

  it('never gives an upload ticket more than one use', () => {
    const store = new AttachmentTicketStore(120);
    const { id } = store.mintUpload('/me/messages/1/attachments', undefined, {
      name: 'a.pdf',
      contentType: 'application/pdf',
    });
    expect(store.redeem(id)).toMatchObject({ purpose: 'upload' });
    expect(store.redeem(id)).toBeUndefined();
  });
});

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
  res.json = () => res;
  res.setHeader = (name: string, value: string) => {
    seen.headers[name] = value;
  };
  return { res, seen, bytes };
}

const okStream = () => ({
  body: new ReadableStream({
    start(c) {
      c.enqueue(new Uint8Array([7]));
      c.close();
    },
  }),
  contentType: 'application/pdf',
  contentLength: 1,
  contentDisposition: null,
});

const authManager = { isOAuthModeEnabled: () => false, getTokenForAccount: async () => 'tok' };

async function hit(handler: ReturnType<typeof createAttachmentHandler>, id: string) {
  const r = ownRes();
  await handler({ method: 'GET', query: { t: id } } as never, r.res as never, (() => {}) as never);
  return r;
}

describe('route: multi-use tickets', () => {
  function setup(downloadStream: () => Promise<unknown>) {
    const store = new AttachmentTicketStore(120);
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () => ({ downloadStream }) as never,
      authManager: authManager as never,
    });
    return { store, handler };
  }

  it('single-use: one fetch, then the uniform 404', async () => {
    const downloadStream = vi.fn(async () => okStream());
    const { store, handler } = setup(downloadStream);
    const { id } = store.mint(TARGET, undefined);
    expect((await hit(handler, id)).seen.status).toBe(200);
    expect((await hit(handler, id)).seen.status).toBe(404);
    expect(downloadStream).toHaveBeenCalledTimes(1);
  });

  it('opt-in: N fetches, then 404', async () => {
    const downloadStream = vi.fn(async () => okStream());
    const { store, handler } = setup(downloadStream);
    const { id } = store.mint(TARGET, undefined, undefined, 3);
    for (let i = 0; i < 3; i += 1) {
      const r = await hit(handler, id);
      expect(r.seen.status).toBe(200);
      expect(Buffer.concat(r.bytes)).toEqual(Buffer.from([7]));
    }
    expect((await hit(handler, id)).seen.status).toBe(404);
    expect(downloadStream).toHaveBeenCalledTimes(3);
  });

  it('a failed Graph fetch still spends one, so a failing target cannot loop', async () => {
    const downloadStream = vi.fn(async () => {
      throw new Error('Graph 404');
    });
    const { store, handler } = setup(downloadStream);
    const { id } = store.mint(TARGET, undefined, undefined, 3);
    for (let i = 0; i < 3; i += 1) expect((await hit(handler, id)).seen.status).toBe(502);
    expect((await hit(handler, id)).seen.status).toBe(404);
    expect(downloadStream).toHaveBeenCalledTimes(3);
  });

  it('a failed fetch releases the hold so the next attempt can run', async () => {
    let fail = true;
    const { store, handler } = setup(async () => {
      if (fail) throw new Error('boom');
      return okStream();
    });
    const { id } = store.mint(TARGET, undefined, undefined, 2);
    expect((await hit(handler, id)).seen.status).toBe(502);
    fail = false;
    expect((await hit(handler, id)).seen.status).toBe(200);
  });

  it('never streams one ticket twice at once, and refuses the loser with the uniform 404', async () => {
    let open!: () => void;
    const gate = new Promise<void>((resolve) => (open = resolve));
    const downloadStream = vi.fn(async () => {
      await gate;
      return okStream();
    });
    const { store, handler } = setup(downloadStream);
    const { id } = store.mint(TARGET, undefined, undefined, 3);

    const first = hit(handler, id);
    await Promise.resolve();
    const loser = await hit(handler, id);
    expect(loser.seen.status).toBe(404);
    expect(downloadStream).toHaveBeenCalledTimes(1);

    open();
    expect((await first).seen.status).toBe(200);
    // The loser spent nothing: two more fetches remain.
    expect((await hit(handler, id)).seen.status).toBe(200);
    expect((await hit(handler, id)).seen.status).toBe(200);
    expect((await hit(handler, id)).seen.status).toBe(404);
  });

  it('releases the hold when the handler throws', async () => {
    const { store, handler } = setup(async () => okStream());
    const { id } = store.mint(TARGET, undefined, undefined, 2);
    const r = ownRes();
    r.res.setHeader = () => {
      throw new Error('header refused');
    };
    await expect(
      handler({ method: 'GET', query: { t: id } } as never, r.res as never, (() => {}) as never)
    ).rejects.toThrow('header refused');
    expect((await hit(handler, id)).seen.status).toBe(200);
  });

  it('releases the hold when the stream aborts mid-body', async () => {
    let abort = true;
    const { store, handler } = setup(async () =>
      abort
        ? {
            ...okStream(),
            body: new ReadableStream({
              start(c) {
                c.enqueue(new Uint8Array([1]));
                c.error(new Error('reset'));
              },
            }),
          }
        : okStream()
    );
    const { id } = store.mint(TARGET, undefined, undefined, 2);
    await hit(handler, id);
    abort = false;
    expect((await hit(handler, id)).seen.status).toBe(200);
  });

  it('a purpose mismatch spends one and releases, answering the uniform 404', async () => {
    const downloadStream = vi.fn(async () => okStream());
    const { store, handler } = setup(downloadStream);
    const upload = createAttachmentUploadHandler({
      store,
      getGraphClient: () => ({}) as never,
      authManager: authManager as never,
    });
    const up = store.mintUpload('/me/messages/1/attachments', undefined, {
      name: 'a',
      contentType: 'text/plain',
    });
    expect((await hit(handler, up.id)).seen.status).toBe(404);

    const dl = store.mint(TARGET, undefined, undefined, 2);
    const r = ownRes();
    await upload(
      { method: 'PUT', query: { t: dl.id }, headers: {} } as never,
      r.res as never,
      (() => {}) as never
    );
    expect(r.seen.status).toBe(404);
    // Released, not stranded: one redemption remains.
    expect((await hit(handler, dl.id)).seen.status).toBe(200);
    expect((await hit(handler, dl.id)).seen.status).toBe(404);
  });
});

describe('mintDownloadUrl count', () => {
  it('stays single-use by default and advertises the count when asked', async () => {
    const store = new AttachmentTicketStore(120);
    configureAttachmentMinting({
      store,
      config: { base: 'http://127.0.0.1:1', key: Buffer.from('k'), keyId: 'k', ttlSeconds: 120 },
    } as never);
    try {
      const body = async (opts?: { maxRedemptions?: number }) =>
        JSON.parse(
          (
            (await mintDownloadUrl(TARGET, undefined, undefined, opts))!.content[0] as {
              text: string;
            }
          ).text
        );
      const single = await body();
      expect(single.singleUse).toBe(true);
      expect(single).not.toHaveProperty('maxFetches');
      const multi = await body({ maxRedemptions: 3 });
      expect(multi).toMatchObject({ singleUse: false, maxFetches: 3 });
    } finally {
      resetAttachmentMinting();
    }
  });
});
