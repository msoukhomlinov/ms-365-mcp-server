import { describe, it, expect, vi } from 'vitest';
import { Writable } from 'node:stream';
import { parseContentLengthHeader } from '../src/graph-client.js';
import { createAttachmentHandler } from '../src/attachment-route.js';
import { AttachmentTicketStore } from '../src/lib/attachment-tickets.js';

/**
 * Regression cover for a defect that served every mail attachment as an empty
 * 200.
 *
 * `Headers.get()` answers `null` for an absent header and `Number(null)` is
 * `0` -- a *finite* number -- so `Number.isFinite(Number(h.get(...)))` turned
 * "Graph stated no length" into "Graph stated zero". Graph does not send
 * `content-length` on `/$value` attachment responses, so that was the normal
 * path, not an edge case. The route then declared `content-length: 0`, Node
 * ended the response after zero bytes, and the sidecar read a clean,
 * well-formed, empty `application/pdf` that no retry logic would question.
 *
 * It shipped because every existing route test mocks `downloadStream` and
 * hands it a truthful `contentLength`. The mock told the truth the real client
 * could not, so nothing exercised the derivation at all. These tests cover
 * both halves independently: the parser, and the route's refusal to re-emit a
 * non-positive length on a body it is about to stream.
 */
describe('parseContentLengthHeader', () => {
  it('answers null for an absent header rather than 0', () => {
    // The whole defect in one assertion.
    expect(parseContentLengthHeader(null)).toBeNull();
    expect(parseContentLengthHeader(undefined)).toBeNull();
  });

  it('reads a plain digit run', () => {
    expect(parseContentLengthHeader('68764')).toBe(68764);
    expect(parseContentLengthHeader('  68764  ')).toBe(68764);
  });

  it('treats a literal 0 as a real statement by the upstream', () => {
    // Distinct from "absent". Refusing to re-emit it is the consumer's job.
    expect(parseContentLengthHeader('0')).toBe(0);
  });

  it('refuses anything that is not a bare digit run', () => {
    for (const bad of ['', '   ', 'abc', '-5', '1.5', '1e3', '+7', '12abc', '0x10']) {
      expect(parseContentLengthHeader(bad), `expected null for ${JSON.stringify(bad)}`).toBeNull();
    }
  });

  it('refuses a value past MAX_SAFE_INTEGER instead of returning an imprecise number', () => {
    expect(parseContentLengthHeader('9007199254740993')).toBeNull();
  });
});

describe('the attachment route never declares a length it is not about to send', () => {
  const authManager = { isOAuthModeEnabled: () => false, getTokenForAccount: async () => 'tok' };

  function harness(contentLength: number | null, bytes: number[]) {
    const store = new AttachmentTicketStore(120);
    const sent: { status?: number; body?: unknown; headers: Record<string, string> } = {
      headers: {},
    };
    const written: Buffer[] = [];
    // A real Writable, not a recorder: the route awaits `pipeline(...)`, which
    // only settles against a genuine stream. Same reason the sibling suite
    // does this.
    const res = new Writable({
      write(chunk: Buffer, _enc, cb) {
        written.push(Buffer.from(chunk));
        cb();
      },
    }) as Writable & Record<string, unknown>;
    res.status = (s: number) => ((sent.status = s), res);
    res.type = () => res;
    res.send = (body: unknown) => ((sent.body = body), res);
    res.setHeader = (k: string, v: string) => {
      sent.headers[k.toLowerCase()] = v;
    };
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () =>
        ({
          downloadStream: vi.fn(async () => ({
            body: new ReadableStream({
              start(c) {
                if (bytes.length) c.enqueue(new Uint8Array(bytes));
                c.close();
              },
            }),
            contentType: 'application/pdf',
            contentLength,
            contentDisposition: null,
          })),
        }) as never,
      authManager: authManager as never,
    });
    return { store, handler, sent, written, res };
  }

  it('omits content-length when Graph stated none, and still sends the bytes', async () => {
    // The live case: Graph sends no content-length on /$value.
    const { store, handler, sent, written, res } = harness(null, [1, 2, 3]);
    const { id } = store.mint('/me/messages/1/attachments/2/$value', undefined);
    await handler({ query: { t: id } } as never, res as never, (() => {}) as never);

    expect(sent.status).toBe(200);
    expect(sent.headers['content-length']).toBeUndefined();
    expect(Buffer.concat(written)).toEqual(Buffer.from([1, 2, 3]));
  });

  it('omits a zero length rather than truncating a non-empty body', async () => {
    // Even if something upstream genuinely says 0, declaring it on a body we
    // are about to stream produces the silent empty 200.
    const { store, handler, sent, written, res } = harness(0, [1, 2, 3]);
    const { id } = store.mint('/t', undefined);
    await handler({ query: { t: id } } as never, res as never, (() => {}) as never);

    expect(sent.headers['content-length']).toBeUndefined();
    expect(Buffer.concat(written)).toEqual(Buffer.from([1, 2, 3]));
  });

  it('declares a real positive length', async () => {
    const { store, handler, sent, res } = harness(3, [1, 2, 3]);
    const { id } = store.mint('/t', undefined);
    await handler({ query: { t: id } } as never, res as never, (() => {}) as never);

    expect(sent.headers['content-length']).toBe('3');
  });

  it('drops a nonsensical length instead of trusting it', async () => {
    for (const bogus of [-5, 1.5]) {
      const { store, handler, sent, written, res } = harness(bogus, [1, 2, 3]);
      const { id } = store.mint('/t', undefined);
      await handler({ query: { t: id } } as never, res as never, (() => {}) as never);
      // Assert the success path too. An error response also carries no
      // content-length, so checking only the header would pass on a 502 and
      // prove nothing.
      expect(sent.status, `for ${bogus}`).toBe(200);
      expect(sent.headers['content-length'], `for ${bogus}`).toBeUndefined();
      expect(Buffer.concat(written), `for ${bogus}`).toEqual(Buffer.from([1, 2, 3]));
    }
  });
});
