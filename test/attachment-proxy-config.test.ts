/**
 * `--attachment-proxy`: the flag, the environment variable, and the two things
 * it has to imply.
 *
 * This task deliberately stops short of the tool the flag exists to serve.
 * `read-document` is registered, and mapped into `FLAG_UNIVERSAL_UTILITY_TOOLS`, by a later
 * task that adds the mapping entry and the tool together, in the same change -- so a preset can
 * never claim a tool that no build of the server actually has. Until then `attachmentProxy` is a
 * real field on `PresetToolOptions` that gates nothing, and the tests below prove exactly that:
 * the flag, the implication, and the preset-pattern plumbing all work, while the preset contents
 * stay byte-for-byte what they were before this flag existed.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { getCategoryPattern, TOOL_CATEGORIES } from '../src/tool-categories.js';

const commanderMocks = vi.hoisted(() => {
  const mockCommand = {
    name: vi.fn().mockReturnThis(),
    description: vi.fn().mockReturnThis(),
    version: vi.fn().mockReturnThis(),
    option: vi.fn().mockReturnThis(),
    addOption: vi.fn().mockReturnThis(),
    parse: vi.fn(),
    opts: vi.fn().mockReturnValue({}),
  };
  return { mockCommand };
});

vi.mock('commander', () => {
  class MockOption {
    constructor(
      public flags: string,
      public description: string
    ) {}
    hideHelp() {
      return this;
    }
  }
  return {
    Command: vi.fn(function () {
      return commanderMocks.mockCommand;
    }),
    Option: MockOption,
  };
});

const exitSpy = vi.spyOn(process, 'exit').mockImplementation((() => {}) as never);
vi.spyOn(console, 'error').mockImplementation(() => {});
vi.spyOn(console, 'warn').mockImplementation(() => {});

import { parseArgs } from '../src/cli.js';

describe('--attachment-proxy / MS365_MCP_ATTACHMENT_PROXY', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    delete process.env.MS365_MCP_ATTACHMENT_PROXY;
    commanderMocks.mockCommand.opts.mockReturnValue({});
  });

  afterEach(() => {
    delete process.env.MS365_MCP_ATTACHMENT_PROXY;
  });

  it('implies --enable-attachment-urls, because the proxy fetches from that listener', () => {
    commanderMocks.mockCommand.opts.mockReturnValue({
      attachmentProxy: 'http://docglean:8080/mcp',
    });
    const options = parseArgs();
    expect(options.attachmentProxy).toBe('http://docglean:8080/mcp');
    expect(options.enableAttachmentUrls).toBe(true);
  });

  it('reads the environment when the flag is absent', () => {
    process.env.MS365_MCP_ATTACHMENT_PROXY = 'http://docglean:8080/mcp';
    const options = parseArgs();
    expect(options.attachmentProxy).toBe('http://docglean:8080/mcp');
    expect(options.enableAttachmentUrls).toBe(true);
  });

  it('lets the command line win over the environment', () => {
    process.env.MS365_MCP_ATTACHMENT_PROXY = 'http://from-env:8080/mcp';
    commanderMocks.mockCommand.opts.mockReturnValue({
      attachmentProxy: 'http://from-argv:8080/mcp',
    });
    expect(parseArgs().attachmentProxy).toBe('http://from-argv:8080/mcp');
  });

  it('refuses a value that is not an absolute http(s) URL, without enabling anything', () => {
    for (const bad of ['docglean:8080/mcp', '/mcp', '', 'ftp://docglean/mcp', 'not a url']) {
      vi.clearAllMocks();
      commanderMocks.mockCommand.opts.mockReturnValue({ attachmentProxy: bad });
      const options = parseArgs();
      expect(exitSpy, `expected exit(1) for ${JSON.stringify(bad)}`).toHaveBeenCalledWith(1);
      // The dangerous half: a refused proxy must not leave the minting feature
      // switched on behind it.
      expect(options.enableAttachmentUrls).toBeFalsy();
    }
  });

  it('leaves everything alone when neither flag nor env is present', () => {
    const options = parseArgs();
    expect(options.attachmentProxy).toBeUndefined();
    expect(options.enableAttachmentUrls).toBeUndefined();
    expect(exitSpy).not.toHaveBeenCalled();
  });

  it('computes the preset pattern with the flag visible, even from the environment', () => {
    // The ordering bug this guards against: if the env var were read after the
    // --preset block runs, the enableAttachmentUrls implication below would not
    // be visible yet, and this pattern would come back without get-download-url
    // in it -- exactly the "enabled, validated, unreachable" failure
    // --enable-attachment-urls already shipped once. get-download-url is the
    // real, already-gated tool this implication actually reaches today;
    // read-document is not (see the file header).
    process.env.MS365_MCP_ATTACHMENT_PROXY = 'http://docglean:8080/mcp';
    commanderMocks.mockCommand.opts.mockReturnValue({ preset: 'mail,calendar,tasks,contacts' });
    const options = parseArgs();
    expect(new RegExp(options.enabledTools as string).test('get-download-url')).toBe(true);
  });
});

describe('attachmentProxy is on PresetToolOptions but gates nothing yet', () => {
  // The other half of the invariant this whole feature exists to hold: not just
  // "unset changes nothing," but "present and even set to true changes nothing
  // right now," because no FLAG_UNIVERSAL_UTILITY_TOOLS entry answers to this
  // key yet. A test proving that matters as much as one proving the eventual
  // gated case works -- and once a later task adds the mapping entry together
  // with the tool it gates, this exact test is what will need to change, which
  // is the point: the change becomes visible instead of silent.
  it.each(Object.keys(TOOL_CATEGORIES).filter((name) => name !== 'all'))(
    'attachmentProxy leaves preset %s byte-for-byte unchanged',
    (preset) => {
      const off = getCategoryPattern(preset, {})!;
      const on = getCategoryPattern(preset, { attachmentProxy: true })!;
      expect(on.source).toBe(off.source);
    }
  );
});
