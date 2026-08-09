/**
 * Regression cover for the confirmed live defect: `read-document` converts a
 * PDF correctly but silently mis-converts (or mangles) a nested-message
 * attachment, because `attachment-route.ts` forwards whatever Content-Type
 * header Graph's `/$value` response happens to carry -- and `graph-client.ts`
 * defaults that header to `application/octet-stream` whenever Graph sends
 * none at all. A generic subtype matches no format in docglean's table, so
 * docglean falls through to a plain-text path instead of routing to the
 * converter that actually exists for the real format.
 *
 * Verified live against the deployed mailbox (see the SDD content-type
 * report): a mail `itemAttachment` (a nested forwarded message) came back
 * with a `null` Content-Type in Graph's own attachment metadata too, so the
 * probe cannot recover a real value there by reading metadata alone -- the
 * fix additionally falls back to Graph's documented `itemAttachment.$value`
 * contract (always the RFC 5322 source) when the metadata's own probe found
 * nothing and the concrete attachment type is a message item.
 *
 * These tests exercise the route's precedence rule directly, with
 * `downloadStream` mocked -- the same harness `attachment-content-length.test.ts`
 * uses -- so the fetch itself is out of scope; the minting side that produces
 * a ticket's `probedContentType` is covered separately, in
 * `attachment-proxy-read-document.test.ts` (read-document) and
 * `attachment-mint-identity.test.ts` (get-download-url).
 */
import { describe, it, expect } from 'vitest';
import { Writable } from 'node:stream';
import { createAttachmentHandler } from '../src/attachment-route.js';
import { AttachmentTicketStore } from '../src/lib/attachment-tickets.js';

describe('the attachment route resolves Content-Type by precedence, not by forwarding blindly', () => {
  const authManager = { isOAuthModeEnabled: () => false, getTokenForAccount: async () => 'tok' };

  function harness(streamContentType: string, streamContentDisposition: string | null) {
    const store = new AttachmentTicketStore(120);
    const sent: { status?: number; headers: Record<string, string> } = { headers: {} };
    const written: Buffer[] = [];
    const res = new Writable({
      write(chunk: Buffer, _enc, cb) {
        written.push(Buffer.from(chunk));
        cb();
      },
    }) as Writable & Record<string, unknown>;
    res.status = (s: number) => ((sent.status = s), res);
    res.type = () => res;
    res.send = () => res;
    res.setHeader = (k: string, v: string) => {
      sent.headers[k.toLowerCase()] = v;
    };
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () =>
        ({
          downloadStream: async () => ({
            body: new ReadableStream({
              start(c) {
                c.enqueue(new Uint8Array([1, 2, 3]));
                c.close();
              },
            }),
            contentType: streamContentType,
            contentLength: 3,
            contentDisposition: streamContentDisposition,
          }),
        }) as never,
      authManager: authManager as never,
    });
    return { store, handler, sent, written, res };
  }

  it('serves the probed type when the stream carries the generic Graph default', async () => {
    const { store, handler, sent, res } = harness('application/octet-stream', null);
    const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
      contentType: 'message/rfc822',
      name: 'Katusha',
    });
    await handler({ query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-type']).toBe('message/rfc822');
  });

  it('serves the probed type when the stream Content-Type is empty', async () => {
    const { store, handler, sent, res } = harness('', null);
    const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
      contentType: 'message/rfc822',
      name: null,
    });
    await handler({ query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-type']).toBe('message/rfc822');
  });

  it('serves the probed type when the stream carries application/binary', async () => {
    const { store, handler, sent, res } = harness('application/binary', null);
    const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
      contentType: 'application/pdf',
      name: null,
    });
    await handler({ query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-type']).toBe('application/pdf');
  });

  it('serves the STREAM type when it is specific, even if a probed type is also present', async () => {
    // Specific-but-wrong-sounding beats a probe that may be stale: the probe
    // ran at mint time, the stream is what Graph is answering right now.
    const { store, handler, sent, res } = harness('application/pdf', null);
    const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
      contentType: 'message/rfc822',
      name: 'irrelevant.eml',
    });
    await handler({ query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-type']).toBe('application/pdf');
  });

  it('keeps the generic default when neither axis has anything specific', async () => {
    const { store, handler, sent, res } = harness('application/octet-stream', null);
    const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined);
    await handler({ query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-type']).toBe('application/octet-stream');
  });

  it('treats a charset-qualified generic type as generic too', async () => {
    const { store, handler, sent, res } = harness('application/octet-stream; charset=binary', null);
    const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
      contentType: 'message/rfc822',
      name: null,
    });
    await handler({ query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-type']).toBe('message/rfc822');
  });

  it('leaves content-disposition alone when Graph already sent one', async () => {
    const { store, handler, sent, res } = harness(
      'application/octet-stream',
      'attachment; filename="x.bin"'
    );
    const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
      contentType: 'message/rfc822',
      name: 'Katusha',
    });
    await handler({ query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-disposition']).toBe('attachment; filename="x.bin"');
  });

  it('falls back to the probed name for content-disposition when Graph sent none', async () => {
    // The cheaper second axis: docglean's own format resolver also reads the
    // extension off Content-Disposition's filename, so handing it a real name
    // when Graph did not is a second, independent way to route correctly.
    const { store, handler, sent, res } = harness('application/octet-stream', null);
    const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
      contentType: null,
      name: 'report.pdf',
    });
    await handler({ query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-disposition']).toBe('attachment; filename="report.pdf"');
  });

  it('sanitizes a probed name before putting it in a header', async () => {
    const { store, handler, sent, res } = harness('application/octet-stream', null);
    const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
      contentType: null,
      name: 'evil"\r\nX-Injected: yes\r\n.pdf',
    });
    await handler({ query: { t: id } } as never, res as never, (() => {}) as never);
    const value = sent.headers['content-disposition']!;
    // No raw CR/LF anywhere -- that is the actual header-injection vector.
    expect(value).not.toMatch(/[\r\n]/);
    // Exactly the two delimiter quotes around `filename=...`, none smuggled
    // in from the name itself.
    expect(value.match(/"/g)).toHaveLength(2);
    expect(value).toBe('attachment; filename="evilX-Injected: yes.pdf"');
    expect(Object.keys(sent.headers)).not.toContain('x-injected');
  });

  it('defaults content-disposition to a bare attachment when neither Graph nor the probe named the file', async () => {
    const { store, handler, sent, res } = harness('application/octet-stream', null);
    const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined);
    await handler({ query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-disposition']).toBe('attachment');
  });
});
