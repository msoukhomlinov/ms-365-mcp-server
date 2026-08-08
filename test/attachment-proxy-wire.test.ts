import { describe, it, expect } from 'vitest';
import { toWireArguments } from '../src/lib/attachment-proxy.js';

/**
 * The TS interface is camelCase and the wire is snake_case. A `maxChars` that
 * never becomes `max_chars` is not an error anywhere: docglean applies its own
 * default cap and returns a perfectly good, wrongly-sized document.
 */
describe('toWireArguments', () => {
  it('renames maxChars to max_chars and leaves the rest alone', () => {
    expect(
      toWireArguments({ uri: 'https://p/a?t=1', pages: '1-3', offset: 40, maxChars: 20000 })
    ).toEqual({
      uri: 'https://p/a?t=1',
      pages: '1-3',
      offset: 40,
      max_chars: 20000,
    });
  });

  it('omits every optional parameter that was not given, rather than sending null', () => {
    const args = toWireArguments({ uri: 'https://p/a?t=1' });
    expect(args).toEqual({ uri: 'https://p/a?t=1' });
    // Absent, not present-and-undefined: JSON.stringify drops undefined, but
    // `in` is what a reviewer can check, and a null here is invalid_max_chars.
    expect('max_chars' in args).toBe(false);
    expect('offset' in args).toBe(false);
    expect('pages' in args).toBe(false);
  });

  it('keeps a zero offset, which is a real value and not an absent one', () => {
    expect(toWireArguments({ uri: 'u', offset: 0 })).toEqual({ uri: 'u', offset: 0 });
  });

  it('never invents an auth_profile', () => {
    expect('auth_profile' in toWireArguments({ uri: 'u' })).toBe(false);
  });
});
