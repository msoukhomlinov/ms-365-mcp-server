import { describe, it, expect } from 'vitest';
import { scrubByteFields } from '../src/lib/response-scrubber.js';

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
});
