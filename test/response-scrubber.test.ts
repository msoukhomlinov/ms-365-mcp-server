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
});
