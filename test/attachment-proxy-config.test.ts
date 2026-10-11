/**
 * `--attachment-proxy`: the flag, the environment variable, and the two things
 * it has to imply.
 *
 * `getCombinedPresetPattern` adds `read-document` under `attachmentProxy`, to every preset that
 * carries the byte readers. The tests below prove the flag, the implication, and the
 * preset-pattern plumbing all work, and that the ONLY preset change the flag makes is adding
 * read-document -- everything else stays what it was before this flag existed. The registration-layer proof (the three byte tools actually
 * disappearing, not just the pattern gaining a name) lives in
 * `test/attachment-proxy-registration.test.ts`.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { getCombinedPresetPattern, TOOL_CATEGORIES } from '../src/tool-categories.js';
import { UTILITY_TOOLS } from '../src/graph-tools.js';
import { readFileSync } from 'node:fs';

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

  it.each([
    ['--toon', { toon: true }, undefined],
    ['MS365_MCP_OUTPUT_FORMAT=toon', {}, 'toon'],
  ])('refuses %s together with the proxy, which cannot scrub TOON', (_name, opts, env) => {
    if (env) process.env.MS365_MCP_OUTPUT_FORMAT = env;
    try {
      commanderMocks.mockCommand.opts.mockReturnValue({
        attachmentProxy: 'http://docglean:8080/mcp',
        ...opts,
      });
      parseArgs();
      expect(exitSpy).toHaveBeenCalledWith(1);
    } finally {
      delete process.env.MS365_MCP_OUTPUT_FORMAT;
    }
  });

  it('still allows --toon without the proxy', () => {
    commanderMocks.mockCommand.opts.mockReturnValue({ toon: true });
    expect(parseArgs().toon).toBe(true);
    expect(exitSpy).not.toHaveBeenCalled();
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
    // (or, now, read-document) in it -- exactly the "enabled, validated,
    // unreachable" failure --enable-attachment-urls already shipped once.
    process.env.MS365_MCP_ATTACHMENT_PROXY = 'http://docglean:8080/mcp';
    commanderMocks.mockCommand.opts.mockReturnValue({
      preset: 'mail,calendar,tasks,contacts',
      http: '3000',
    });
    const options = parseArgs();
    expect(new RegExp(options.enabledTools as string).test('get-download-url')).toBe(true);
    expect(new RegExp(options.enabledTools as string).test('read-document')).toBe(true);
  });
});

describe('attachmentProxy widens every preset by exactly one tool: read-document', () => {
  const endpointNames: string[] = (
    JSON.parse(readFileSync(new URL('../src/endpoints.json', import.meta.url), 'utf8')) as Array<{
      toolName: string;
    }>
  ).map((e) => e.toolName);
  const allNames = [...new Set([...endpointNames, ...UTILITY_TOOLS.map((u) => u.name)])];
  const matched = (preset: string, attachmentProxy: boolean) => {
    const re = new RegExp(getCombinedPresetPattern([preset], { attachmentProxy }));
    return new Set(allNames.filter((n) => re.test(n)));
  };

  it.each(Object.keys(TOOL_CATEGORIES).filter((name) => name !== 'all' && name !== 'teams-write'))(
    'attachmentProxy adds read-document to preset %s and changes nothing else',
    (preset) => {
      expect(matched(preset, true)).toEqual(new Set([...matched(preset, false), 'read-document']));
    }
  );

  it('adds nothing to the send-only preset, which carries no byte readers', () => {
    expect(matched('teams-write', true)).toEqual(matched('teams-write', false));
  });
});
