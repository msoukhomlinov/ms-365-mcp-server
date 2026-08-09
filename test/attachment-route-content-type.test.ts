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
 * report): Graph's own attachment metadata sometimes populates `contentType`
 * directly for a mail `itemAttachment` (a nested forwarded message) --
 * `message/rfc822`, read verbatim, no inference -- and sometimes leaves it
 * `null` for a different itemAttachment in the same mailbox. The probe used
 * by `graph-tools.ts` relies only on the former (verified) case; the route
 * tested here is agnostic to how the ticket's `probedContentType` was
 * decided and just applies the precedence rule against whatever it carries.
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
import { OutgoingMessage } from 'node:http';
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
    // A throwaway REAL http.OutgoingMessage validates the header the exact
    // way Node's actual response object would -- synchronously, no socket
    // needed for setHeader itself -- rather than a permissive mock that
    // would happily accept a value real Node rejects. This is the harness
    // gap that let the RFC 5987 defect (ERR_INVALID_CHAR on a non-Latin-1
    // probed filename) ship unnoticed: every prior test here used a plain
    // object assignment for setHeader, which cannot throw on anything.
    res.setHeader = (k: string, v: string) => {
      new OutgoingMessage().setHeader(k, v);
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
    // The ASCII-safe filename= param has exactly its two delimiter quotes,
    // none smuggled in from the name itself.
    const asciiParam = /filename="([^]*?)"/.exec(value)!;
    expect(asciiParam[1]).toBe('evilX-Injected: yes.pdf');
    // The raw CR/LF is outside HEADER_SAFE_CHAR, so this name also gets an
    // RFC 5987 extended form -- percent-encoded, so still injection-safe.
    expect(value).toContain("; filename*=UTF-8''");
    expect(Object.keys(sent.headers)).not.toContain('x-injected');
  });

  it('defaults content-disposition to a bare attachment when neither Graph nor the probe named the file', async () => {
    const { store, handler, sent, res } = harness('application/octet-stream', null);
    const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined);
    await handler({ query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-disposition']).toBe('attachment');
  });

  /**
   * Regression cover for a defect this PR itself introduced: the probed-name
   * Content-Disposition fallback above is new, and Node's `setHeader` rejects
   * ANY code point past U+00FF in a header value (`ERR_INVALID_CHAR`) --
   * verified directly against a real `http.OutgoingMessage` (see this file's
   * harness). Before this fallback existed, Graph sending no
   * Content-Disposition just meant no header; after it, a non-Latin-1 probed
   * `name` (a perfectly ordinary email attachment filename, e.g. `报告.pdf`)
   * made every redemption of that ticket 500 instead of streaming the bytes
   * -- spending one of the caller's 3 fetches for nothing.
   */
  describe('content-disposition survives a non-Latin-1 probed name', () => {
    it('leaves a pure-ASCII name exactly as before (no RFC 5987 extension)', async () => {
      const { store, handler, sent, res } = harness('application/octet-stream', null);
      const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
        contentType: null,
        name: 'report.pdf',
      });
      await expect(
        handler({ query: { t: id } } as never, res as never, (() => {}) as never)
      ).resolves.not.toThrow();
      expect(sent.headers['content-disposition']).toBe('attachment; filename="report.pdf"');
    });

    it('leaves a Latin-1-representable name unextended too (café.pdf is valid in a raw header value)', async () => {
      // Node's own header-value validator accepts the raw Latin-1 supplement
      // (0x80-0xFF) directly -- HTTP header values are historically
      // ISO-8859-1 -- so a name like this never needed RFC 5987 at all.
      const { store, handler, sent, res } = harness('application/octet-stream', null);
      const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
        contentType: null,
        name: 'café.pdf',
      });
      await expect(
        handler({ query: { t: id } } as never, res as never, (() => {}) as never)
      ).resolves.not.toThrow();
      const value = sent.headers['content-disposition']!;
      expect(value).toBe('attachment; filename="café.pdf"');
      expect(value).not.toContain('filename*=');
    });

    it('does not throw ERR_INVALID_CHAR for a non-Latin-1 name, and emits a well-formed RFC 5987 pair', async () => {
      const { store, handler, sent, res } = harness('application/octet-stream', null);
      const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
        contentType: null,
        name: '报告.pdf',
      });
      await expect(
        handler({ query: { t: id } } as never, res as never, (() => {}) as never)
      ).resolves.not.toThrow();
      const value = sent.headers['content-disposition']!;
      // ASCII-safe fallback recovers at least the extension, so an agent or
      // client reading only `filename=` (the RFC 5987 fallback rule) still
      // gets a usable, format-bearing name rather than an empty one.
      expect(value).toContain('filename="attachment.pdf"');
      expect(value).toContain("filename*=UTF-8''%E6%8A%A5%E5%91%8A.pdf");
      // The RFC 5987 grammar's attr-char set excludes ' ( ) * -- encodeURIComponent
      // alone leaves those unescaped, so a correct encoder must not either.
      expect(value).not.toMatch(/filename\*=UTF-8''[^;]*['()*]/);
    });

    it('recovers gracefully when the name is entirely non-ASCII with no extension to fall back to', async () => {
      const { store, handler, sent, res } = harness('application/octet-stream', null);
      const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
        contentType: null,
        name: '报告',
      });
      await expect(
        handler({ query: { t: id } } as never, res as never, (() => {}) as never)
      ).resolves.not.toThrow();
      const value = sent.headers['content-disposition']!;
      expect(value).toContain('filename="attachment"');
      expect(value).toContain("filename*=UTF-8''%E6%8A%A5%E5%91%8A");
    });

    it('does not throw for a quote/backslash mixed with non-Latin-1 characters', async () => {
      const { store, handler, sent, res } = harness('application/octet-stream', null);
      const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
        contentType: null,
        name: '报"告\\.pdf',
      });
      await expect(
        handler({ query: { t: id } } as never, res as never, (() => {}) as never)
      ).resolves.not.toThrow();
      const value = sent.headers['content-disposition']!;
      expect(value).not.toMatch(/[\r\n]/);
      // Exactly the two delimiter quotes -- none smuggled in from the name.
      const asciiParam = /filename="([^]*?)"/.exec(value)!;
      expect(asciiParam[1]).not.toMatch(/["\\]/);
    });

    it('does not throw for a very long non-ASCII name', async () => {
      const longName = '报'.repeat(300) + '.pdf';
      const { store, handler, sent, res } = harness('application/octet-stream', null);
      const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
        contentType: null,
        name: longName,
      });
      await expect(
        handler({ query: { t: id } } as never, res as never, (() => {}) as never)
      ).resolves.not.toThrow();
      expect(sent.headers['content-disposition']).toContain('filename="attachment.pdf"');
    });
  });

  describe('content-type from a probe is guarded the same way', () => {
    it('falls back to the stream type rather than throwing when the probed content-type is unusable', async () => {
      // A metadata field, not an HTTP header Graph itself sent -- so nothing
      // upstream already constrained it to header-safe characters the way
      // stream.contentType is. Same class of defect as the filename one
      // above, same fix standard: never let untrusted probe text reach
      // setHeader unguarded.
      const { store, handler, sent, res } = harness('application/pdf', null);
      const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
        contentType: 'application/pdf; name=报告.pdf',
        name: null,
      });
      await expect(
        handler({ query: { t: id } } as never, res as never, (() => {}) as never)
      ).resolves.not.toThrow();
      expect(sent.headers['content-type']).toBe('application/pdf');
    });

    it('falls back to the generic default when the probed content-type is unusable AND the stream is also generic', async () => {
      const { store, handler, sent, res } = harness('application/octet-stream', null);
      const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined, undefined, {
        contentType: 'message/rfc822; x=报告',
        name: null,
      });
      await expect(
        handler({ query: { t: id } } as never, res as never, (() => {}) as never)
      ).resolves.not.toThrow();
      expect(sent.headers['content-type']).toBe('application/octet-stream');
    });
  });
});
