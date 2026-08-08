import { describe, it, expect } from 'vitest';
import { randomBytes } from 'node:crypto';
import { BASE64_STRIP_THRESHOLD, scrubByteFields } from '../src/lib/response-scrubber.js';

/** Valid standard base64, `chars` long. `chars` must be a multiple of 4. */
function base64OfLength(chars: number): string {
  return randomBytes((chars / 4) * 3).toString('base64');
}

describe('scrubByteFields: the contentBytes name rule', () => {
  it('strips a field named contentBytes and names read-document in its place', () => {
    const result = scrubByteFields({ name: 'a.pdf', contentBytes: 'QUJDRA==' });
    expect(result.value).toEqual({
      name: 'a.pdf',
      contentBytes: '<stripped: 4 bytes, use read-document>',
    });
  });

  it('reports the strip rather than only performing it', () => {
    const result = scrubByteFields({ contentBytes: 'QUJDRA==' });
    expect(result.stripped).toEqual([{ path: '$.contentBytes', field: 'contentBytes', bytes: 4 }]);
  });

  it('leaves a null contentBytes alone, as reference attachments carry', () => {
    const input = { name: 'link.url', contentBytes: null };
    const result = scrubByteFields(input);
    expect(result.value).toBe(input);
    expect(result.stripped).toEqual([]);
  });

  it('returns the input by reference when there is nothing to strip', () => {
    const input = { id: 'AAMk', subject: 'hello', body: { content: '<p>hi</p>' } };
    const result = scrubByteFields(input);
    expect(result.value).toBe(input);
    expect(result.stripped).toEqual([]);
  });

  it('does not mutate the input', () => {
    const input = { contentBytes: 'QUJDRA==' };
    scrubByteFields(input);
    expect(input.contentBytes).toBe('QUJDRA==');
  });

  it('strips a non-base64-shaped contentBytes and reports its own UTF-8 length', () => {
    // The name rule does not care about shape: 'not-base64!' is not valid
    // base64 (11 chars, not a multiple of 4), so there is nothing to decode.
    // The bytes removed are the string's own UTF-8 bytes -- 11 of them.
    const result = scrubByteFields({ contentBytes: 'not-base64!' });
    expect(result.value).toEqual({
      contentBytes: '<stripped: 11 bytes, use read-document>',
    });
    expect(result.stripped).toEqual([{ path: '$.contentBytes', field: 'contentBytes', bytes: 11 }]);
  });
});

describe('scrubByteFields: nested objects and arrays', () => {
  it('strips through nested object depth', () => {
    const result = scrubByteFields({
      value: { message: { attachment: { name: 'a.pdf', contentBytes: 'QUJDRA==' } } },
    });
    expect(result.value).toEqual({
      value: {
        message: {
          attachment: { name: 'a.pdf', contentBytes: '<stripped: 4 bytes, use read-document>' },
        },
      },
    });
    expect(result.stripped).toEqual([
      { path: '$.value.message.attachment.contentBytes', field: 'contentBytes', bytes: 4 },
    ]);
  });

  it('strips every attachment in an array of attachments', () => {
    const result = scrubByteFields({
      value: [
        { id: 'm1', attachments: [{ contentBytes: 'QUJDRA==' }, { contentBytes: 'RUZHSA==' }] },
        { id: 'm2', attachments: [{ contentBytes: 'SUpLTA==' }] },
      ],
    });
    expect(result.value).toEqual({
      value: [
        {
          id: 'm1',
          attachments: [
            { contentBytes: '<stripped: 4 bytes, use read-document>' },
            { contentBytes: '<stripped: 4 bytes, use read-document>' },
          ],
        },
        { id: 'm2', attachments: [{ contentBytes: '<stripped: 4 bytes, use read-document>' }] },
      ],
    });
  });

  it('reports a path per strip that locates it inside the response', () => {
    const result = scrubByteFields({
      value: [
        { id: 'm1', attachments: [{ contentBytes: 'QUJDRA==' }, { contentBytes: 'RUZHSA==' }] },
        { id: 'm2', attachments: [{ contentBytes: 'SUpLTA==' }] },
      ],
    });
    expect(result.stripped).toEqual([
      { path: '$.value[0].attachments[0].contentBytes', field: 'contentBytes', bytes: 4 },
      { path: '$.value[0].attachments[1].contentBytes', field: 'contentBytes', bytes: 4 },
      { path: '$.value[1].attachments[0].contentBytes', field: 'contentBytes', bytes: 4 },
    ]);
  });

  it('keeps untouched siblings by reference while replacing the payload', () => {
    const clean = { id: 'm2', subject: 'no attachments' };
    const input = { value: [{ id: 'm1', contentBytes: 'QUJDRA==' }, clean] };
    const result = scrubByteFields(input) as { value: { value: unknown[] } };
    expect(result.value.value[1]).toBe(clean);
  });

  it('returns a fully clean array by reference, preserving the structural-sharing invariant', () => {
    const input = { value: [{ id: 'm1' }, { id: 'm2' }] };
    const result = scrubByteFields(input);
    expect(result.value).toBe(input);
    expect(result.value.value).toBe(input.value);
    expect(result.stripped).toEqual([]);
  });

  it('does not mutate the input array', () => {
    const original = [
      { id: 'm1', subject: 'hello' },
      { id: 'm2', subject: 'world' },
    ];
    const input = { value: original };
    scrubByteFields(input);
    expect(input.value).toBe(original);
  });
});

describe('scrubByteFields: the base64 shape rule', () => {
  it('strips a long base64 string in a field Graph does not call contentBytes', () => {
    // The rule-2 test. A tool upstream adds tomorrow can return bytes under any
    // name; the name list cannot be kept complete, so shape has to carry it.
    const payload = base64OfLength(8192);
    const result = scrubByteFields({ report: { data: payload } });
    expect(result.value).toEqual({ report: { data: '<stripped: 6144 bytes, use read-document>' } });
    expect(result.stripped).toEqual([{ path: '$.report.data', field: 'data', bytes: 6144 }]);
  });

  it('holds the 4,096 floor: 4096 chars kept, the next valid base64 length stripped', () => {
    expect(BASE64_STRIP_THRESHOLD).toBe(4096);
    const atFloor = base64OfLength(4096);
    expect(scrubByteFields({ data: atFloor }).stripped).toEqual([]);

    // 4097 is the next character count, but base64 comes in multiples of four,
    // so a 4097-character string is not base64 at all -- asserted below so the
    // floor is not accidentally proved by the wrong rule. 4100 is the next
    // length that can be both over the floor and valid base64.
    const justOver = 'A'.repeat(4097);
    expect(scrubByteFields({ data: justOver }).stripped).toEqual([]);

    const overFloor = base64OfLength(4100);
    expect(scrubByteFields({ data: overFloor }).stripped).toEqual([
      { path: '$.data', field: 'data', bytes: 3075 },
    ]);
  });

  it('leaves a 152-character Graph id untouched', () => {
    const id =
      'AAMkAGI2THVSAAA=AAMkAGI2NmRlNTk3LTk5MTUtNDgxYi1iMzg3LTRkNzE3MjkzZTk5MABGAAAAAAB1ZmJk' +
      'YWQxLTk5MTUtNDgxYi1iMzg3LTRkNzE3MjkzZTk5MAcAdWZiZGFkMS05OTE1LTQ4YjM=';
    expect(id.length).toBe(152);
    const input = { id, subject: 'Re: invoice' };
    const result = scrubByteFields(input);
    expect(result.value).toBe(input);
    expect(result.stripped).toEqual([]);
  });

  it('leaves a long non-base64 string alone', () => {
    const body =
      'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor. '.repeat(
        300
      );
    expect(body.length).toBeGreaterThan(BASE64_STRIP_THRESHOLD);
    const input = { body: { contentType: 'text', content: body } };
    const result = scrubByteFields(input);
    expect(result.value).toBe(input);
    expect(result.stripped).toEqual([]);
  });

  it('strips a real 291,572-character contentBytes payload', () => {
    // The measured live case: get-mail-message with expand:["attachments"] on
    // one message in Max's mailbox answered 291,572 bytes.
    const pdf = Buffer.concat([Buffer.from('%PDF-1.7\n'), randomBytes(218670)]);
    const contentBytes = pdf.toString('base64');
    expect(contentBytes.length).toBe(291572);

    const result = scrubByteFields({
      id: 'AAMk',
      attachments: [{ name: 'invoice.pdf', contentType: 'application/pdf', contentBytes }],
    });
    expect(result.value).toEqual({
      id: 'AAMk',
      attachments: [
        {
          name: 'invoice.pdf',
          contentType: 'application/pdf',
          contentBytes: '<stripped: 218679 bytes, use read-document>',
        },
      ],
    });
    expect(JSON.stringify(result.value).length).toBeLessThan(300);
  });

  it('strips line-wrapped base64 rather than letting the wrapping hide it', () => {
    const wrapped = (base64OfLength(8192).match(/.{1,76}/g) ?? []).join('\r\n');
    expect(wrapped.length).toBeGreaterThan(8192);
    expect(scrubByteFields({ data: wrapped }).stripped).toEqual([
      { path: '$.data', field: 'data', bytes: 6144 },
    ]);
  });

  it('does not strip a long base64url string: the alphabet restriction is the guard', () => {
    // This test proves the alphabet restriction (excluding - and _) actually works
    // for long strings. Bearer and refresh tokens are plausibly over 4096 chars,
    // unlike Graph IDs (152 chars, which short-circuit on length alone).
    // A regression widening BASE64_PATTERN to accept base64url would pass all
    // other tests but start stripping tokens. This test catches that.
    // The ONLY thing keeping this string unstripped is the base64url alphabet.
    const bytes = randomBytes(3075); // Converts to 4100-char base64
    const standard = bytes.toString('base64');
    expect(standard.length).toBe(4100);
    expect(standard.length % 4).toBe(0);

    // Convert to base64url by replacing + with - and / with _
    const baseurl = standard.replace(/\+/g, '-').replace(/\//g, '_');
    expect(baseurl.length).toBe(4100);
    expect(baseurl).toMatch(/[-_]/); // Confirm it has base64url chars
    expect(baseurl).not.toMatch(/[+/]/); // Confirm no standard base64 chars

    const input = { token: baseurl };
    const result = scrubByteFields(input);
    expect(result.value).toBe(input);
    expect(result.stripped).toEqual([]);
  });
});

describe('scrubByteFields: hostile shapes', () => {
  function nest(depth: number, leaf: unknown): unknown {
    let node: unknown = leaf;
    for (let i = 0; i < depth; i++) node = { child: node };
    return node;
  }

  it('strips a payload nested within the depth cap', () => {
    const result = scrubByteFields(nest(60, { contentBytes: 'QUJDRA==' }));
    expect(result.stripped).toHaveLength(1);
    expect(result.stripped[0].field).toBe('contentBytes');
    expect(result.stripped[0].bytes).toBe(4);
  });

  it('replaces a subtree past the depth cap instead of recursing into it', () => {
    const result = scrubByteFields(nest(200, { contentBytes: 'QUJDRA==' }));
    expect(JSON.stringify(result.value)).toContain(
      '<stripped: nesting deeper than 64 levels, use read-document>'
    );
    // Fails closed: the payload was never reached, and it never reaches a model.
    expect(JSON.stringify(result.value)).not.toContain('QUJDRA==');
    expect(result.stripped).toEqual([
      { path: `$${'.child'.repeat(65)}`, field: 'child', bytes: 0 },
    ]);
  });

  it('terminates on a cyclic object', () => {
    const cyclic: Record<string, unknown> = { name: 'loop' };
    cyclic.self = cyclic;
    const result = scrubByteFields(cyclic);
    // Both keys of the object reached at depth 65 are capped, in key order.
    expect(result.stripped.map((s) => s.field)).toEqual(['name', 'self']);
    expect(result.stripped.every((s) => s.bytes === 0)).toBe(true);
    expect(JSON.stringify(result.value)).toContain('nesting deeper than 64 levels');
  });

  it('passes non-JSON values through untouched', () => {
    const when = new Date('2026-08-08T00:00:00.000Z');
    const input = { when, count: 3, ok: true, missing: null };
    const result = scrubByteFields(input);
    expect(result.value).toBe(input);
    expect((result.value as { when: Date }).when).toBe(when);
  });

  it('scrubs a bare string handed in at the root', () => {
    const result = scrubByteFields(base64OfLength(8192));
    expect(result.value).toBe('<stripped: 6144 bytes, use read-document>');
    expect(result.stripped).toEqual([{ path: '$', field: '$', bytes: 6144 }]);
  });
});
