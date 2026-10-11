/**
 * The redemption route resolves Content-Type by precedence rather than forwarding the
 * response header blindly. Graph answers a mail attachment /$value with the generic default
 * when it has nothing better, which routes a converter to the wrong format. A ticket can carry
 * what Graph's own metadata said at mint time (`probedContentType`, `probedName`); the route
 * prefers the stream's own type unless it is generic, then a header-safe probed type, and falls
 * back to the probed filename for content-disposition. `downloadStream` is mocked, and the
 * response is a real `http.OutgoingMessage` header validator, so a value Node would reject
 * fails here as it would in production.
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
    // A real http.OutgoingMessage validates the value the way Node would: setHeader throws on
    // a code point past U+00FF, which a permissive mock would accept.
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
    const { id } = store.mint(
      '/me/messages/1/attachments/2/$value',
      undefined,
      undefined,
      undefined,
      {
        probedContentType: 'message/rfc822',
        probedName: 'Katusha',
      }
    );
    await handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-type']).toBe('message/rfc822');
  });

  it('serves the probed type when the stream Content-Type is empty', async () => {
    const { store, handler, sent, res } = harness('', null);
    const { id } = store.mint(
      '/me/messages/1/attachments/2/$value',
      undefined,
      undefined,
      undefined,
      {
        probedContentType: 'message/rfc822',
      }
    );
    await handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-type']).toBe('message/rfc822');
  });

  it('serves the probed type when the stream carries application/binary', async () => {
    const { store, handler, sent, res } = harness('application/binary', null);
    const { id } = store.mint(
      '/me/messages/1/attachments/2/$value',
      undefined,
      undefined,
      undefined,
      {
        probedContentType: 'application/pdf',
      }
    );
    await handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-type']).toBe('application/pdf');
  });

  it('serves the STREAM type when it is specific, even if a probed type is also present', async () => {
    // Specific-but-wrong-sounding beats a probe that may be stale: the probe
    // ran at mint time, the stream is what Graph is answering right now.
    const { store, handler, sent, res } = harness('application/pdf', null);
    const { id } = store.mint(
      '/me/messages/1/attachments/2/$value',
      undefined,
      undefined,
      undefined,
      {
        probedContentType: 'message/rfc822',
        probedName: 'irrelevant.eml',
      }
    );
    await handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-type']).toBe('application/pdf');
  });

  it('keeps the generic default when neither axis has anything specific', async () => {
    const { store, handler, sent, res } = harness('application/octet-stream', null);
    const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined);
    await handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-type']).toBe('application/octet-stream');
  });

  it('treats a charset-qualified generic type as generic too', async () => {
    const { store, handler, sent, res } = harness('application/octet-stream; charset=binary', null);
    const { id } = store.mint(
      '/me/messages/1/attachments/2/$value',
      undefined,
      undefined,
      undefined,
      {
        probedContentType: 'message/rfc822',
      }
    );
    await handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-type']).toBe('message/rfc822');
  });

  it('leaves content-disposition alone when Graph already sent one', async () => {
    const { store, handler, sent, res } = harness(
      'application/octet-stream',
      'attachment; filename="x.bin"'
    );
    const { id } = store.mint(
      '/me/messages/1/attachments/2/$value',
      undefined,
      undefined,
      undefined,
      {
        probedContentType: 'message/rfc822',
        probedName: 'Katusha',
      }
    );
    await handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-disposition']).toBe('attachment; filename="x.bin"');
  });

  it('falls back to the probed name for content-disposition when Graph sent none', async () => {
    // A converter that reads a format off the filename extension still routes correctly.
    const { store, handler, sent, res } = harness('application/octet-stream', null);
    const { id } = store.mint(
      '/me/messages/1/attachments/2/$value',
      undefined,
      undefined,
      undefined,
      {
        probedName: 'report.pdf',
      }
    );
    await handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-disposition']).toBe('attachment; filename="report.pdf"');
  });

  it('sanitizes a probed name before putting it in a header', async () => {
    const { store, handler, sent, res } = harness('application/octet-stream', null);
    const { id } = store.mint(
      '/me/messages/1/attachments/2/$value',
      undefined,
      undefined,
      undefined,
      {
        probedName: 'evil"\r\nX-Injected: yes\r\n.pdf',
      }
    );
    await handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never);
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
    await handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never);
    expect(sent.headers['content-disposition']).toBe('attachment');
  });

  // A probed name is Graph metadata, not a header that survived the wire, and setHeader throws
  // on any code point past U+00FF, so a non-Latin-1 name must not turn a redemption into a 500.
  describe('content-disposition survives a non-Latin-1 probed name', () => {
    it('leaves a pure-ASCII name exactly as before (no RFC 5987 extension)', async () => {
      const { store, handler, sent, res } = harness('application/octet-stream', null);
      const { id } = store.mint(
        '/me/messages/1/attachments/2/$value',
        undefined,
        undefined,
        undefined,
        {
          probedName: 'report.pdf',
        }
      );
      await expect(
        handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never)
      ).resolves.not.toThrow();
      expect(sent.headers['content-disposition']).toBe('attachment; filename="report.pdf"');
    });

    it('leaves a Latin-1-representable name unextended too (café.pdf is valid in a raw header value)', async () => {
      // Node's own header-value validator accepts the raw Latin-1 supplement
      // (0x80-0xFF) directly -- HTTP header values are historically
      // ISO-8859-1 -- so a name like this never needed RFC 5987 at all.
      const { store, handler, sent, res } = harness('application/octet-stream', null);
      const { id } = store.mint(
        '/me/messages/1/attachments/2/$value',
        undefined,
        undefined,
        undefined,
        {
          probedName: 'café.pdf',
        }
      );
      await expect(
        handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never)
      ).resolves.not.toThrow();
      const value = sent.headers['content-disposition']!;
      expect(value).toBe('attachment; filename="café.pdf"');
      expect(value).not.toContain('filename*=');
    });

    it('does not throw ERR_INVALID_CHAR for a non-Latin-1 name, and emits a well-formed RFC 5987 pair', async () => {
      const { store, handler, sent, res } = harness('application/octet-stream', null);
      const { id } = store.mint(
        '/me/messages/1/attachments/2/$value',
        undefined,
        undefined,
        undefined,
        {
          probedName: '报告.pdf',
        }
      );
      await expect(
        handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never)
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
      const { id } = store.mint(
        '/me/messages/1/attachments/2/$value',
        undefined,
        undefined,
        undefined,
        {
          probedName: '报告',
        }
      );
      await expect(
        handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never)
      ).resolves.not.toThrow();
      const value = sent.headers['content-disposition']!;
      expect(value).toContain('filename="attachment"');
      expect(value).toContain("filename*=UTF-8''%E6%8A%A5%E5%91%8A");
    });

    it('does not throw for a quote/backslash mixed with non-Latin-1 characters', async () => {
      const { store, handler, sent, res } = harness('application/octet-stream', null);
      const { id } = store.mint(
        '/me/messages/1/attachments/2/$value',
        undefined,
        undefined,
        undefined,
        {
          probedName: '报"告\\.pdf',
        }
      );
      await expect(
        handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never)
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
      const { id } = store.mint(
        '/me/messages/1/attachments/2/$value',
        undefined,
        undefined,
        undefined,
        {
          probedName: longName,
        }
      );
      await expect(
        handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never)
      ).resolves.not.toThrow();
      expect(sent.headers['content-disposition']).toContain('filename="attachment.pdf"');
    });
  });

  describe('content-type from a probe is guarded the same way', () => {
    it('falls back to the stream type rather than throwing when the probed content-type is unusable', async () => {
      const { store, handler, sent, res } = harness('application/pdf', null);
      const { id } = store.mint(
        '/me/messages/1/attachments/2/$value',
        undefined,
        undefined,
        undefined,
        {
          probedContentType: 'application/pdf; name=报告.pdf',
        }
      );
      await expect(
        handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never)
      ).resolves.not.toThrow();
      expect(sent.headers['content-type']).toBe('application/pdf');
    });

    it('falls back to the generic default when the probed content-type is unusable AND the stream is also generic', async () => {
      const { store, handler, sent, res } = harness('application/octet-stream', null);
      const { id } = store.mint(
        '/me/messages/1/attachments/2/$value',
        undefined,
        undefined,
        undefined,
        {
          probedContentType: 'message/rfc822; x=报告',
        }
      );
      await expect(
        handler({ method: 'GET', query: { t: id } } as never, res as never, (() => {}) as never)
      ).resolves.not.toThrow();
      expect(sent.headers['content-type']).toBe('application/octet-stream');
    });
  });
});
