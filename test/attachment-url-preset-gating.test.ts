/**
 * `--enable-attachment-urls` vs. preset membership.
 *
 * `get-download-url` is the only tool in this server that can mint an attachment URL. It used to be
 * scoped to the drive-backed presets on the claim that it resolved drive/SharePoint content "ONLY",
 * which stopped being true the moment the minting flag existed: under the flag the same tool mints
 * for mail and event attachments, meeting recordings, and any other authenticated `/$value` byte
 * endpoint. The scoping outlived the claim, so `--preset mail,calendar,tasks,contacts
 * --enable-attachment-urls` started cleanly, validated the URL base and key, served the route, and
 * registered nothing that could mint. Enabled, validated, unreachable.
 *
 * These tests are written as class rules, not as a guard against that one command line:
 *
 *  - the set of presets the flag serves is **derived** from `MINTABLE_TARGET_PATTERNS` (the
 *    implementation's own answer to "what does this flag act on?") applied to endpoints.json, so a
 *    future mint target widens the required membership automatically instead of leaving a stale
 *    hard-coded list behind;
 *  - on top of that, a catch-all asserts membership in *every* named preset under the flag, which
 *    is the only form that cannot miss an app or a preset added later;
 *  - the startup warning is checked against the same selector registration uses, so it cannot
 *    describe a server that does not exist.
 */

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';
import { createServer, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import {
  FLAG_UNIVERSAL_UTILITY_TOOLS,
  getCategoryPattern,
  getCombinedPresetPattern,
  TOOL_CATEGORIES,
} from '../src/tool-categories.js';
import {
  MINTABLE_TARGET_PATTERNS,
  registerGraphTools,
  selectUtilityTools,
  utilityToolWillRegister,
  UTILITY_TOOLS,
} from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';
import MicrosoftGraphServer from '../src/server.js';
import type AuthManager from '../src/auth.js';
import type { CommandOptions } from '../src/cli.js';
import { resetAttachmentMinting } from '../src/lib/attachment-minting.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
    verbose: vi.fn(),
  },
  enableConsoleLogging: vi.fn(),
}));

import logger from '../src/logger.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const endpoints: Array<{ toolName: string; pathPattern: string; presets?: string[] }> = JSON.parse(
  readFileSync(path.join(__dirname, '..', 'src', 'endpoints.json'), 'utf8')
);

const NAMED_PRESETS = Object.keys(TOOL_CATEGORIES).filter((name) => name !== 'all');
const UTILITY_NAMES = UTILITY_TOOLS.map((utility) => utility.name);
const ALL_TOOL_NAMES = [...new Set([...endpoints.map((e) => e.toolName), ...UTILITY_NAMES])];
const FLAG_UNIVERSAL_NAMES = Object.keys(FLAG_UNIVERSAL_UTILITY_TOOLS);

/** Presets this preset filter would let through, over the whole tool set. */
function toolsMatching(pattern: string | RegExp): string[] {
  const re = new RegExp(typeof pattern === 'string' ? pattern : pattern.source, 'i');
  return ALL_TOOL_NAMES.filter((name) => re.test(name));
}

function presetContains(preset: string, tool: string, attachmentUrls: boolean): boolean {
  const pattern = getCategoryPattern(preset, { attachmentUrls });
  expect(pattern, `no pattern for preset ${preset}`).toBeDefined();
  return new RegExp(pattern!.source, 'i').test(tool);
}

/**
 * The presets whose resources `--enable-attachment-urls` serves, derived rather than listed.
 *
 * `MINTABLE_TARGET_PATTERNS` is the implementation's definition of a target the flag mints for.
 * Applying it to the endpoint paths this server exposes, and collecting those endpoints' presets,
 * answers "which presets contain a resource this flag was built for?" without anyone restating it.
 */
const FLAG_SERVED_PRESETS: string[] = [
  ...new Set(
    endpoints
      .filter((endpoint) => MINTABLE_TARGET_PATTERNS.some((p) => p.test(endpoint.pathPattern)))
      .flatMap((endpoint) => endpoint.presets ?? [])
  ),
].sort();

describe('presets the attachment-URL flag serves (derived from the mint targets)', () => {
  // If the derivation ever collapses to nothing - a renamed field, a narrowed pattern - every rule
  // built on it would pass vacuously. Assert it found something, and that it found the two
  // resource families the flag exists for (mail attachments, meeting recordings) in presets that
  // are not drive-backed, which is the whole failure this file is about.
  it('derives a non-empty set that includes mail- and teams-shaped presets', () => {
    expect(FLAG_SERVED_PRESETS.length).toBeGreaterThan(0);
    expect(FLAG_SERVED_PRESETS).toContain('mail');
    expect(FLAG_SERVED_PRESETS).toContain('teams');
  });

  it.each(FLAG_UNIVERSAL_NAMES)(
    '%s is registered in every preset whose resources the flag serves',
    (tool) => {
      const missing = FLAG_SERVED_PRESETS.filter((preset) => !presetContains(preset, tool, true));
      expect(missing, `${tool} missing from flag-served presets`).toEqual([]);
    }
  );

  // The catch-all the enumerated version cannot be: a mint target has no preset of its own (event
  // attachments have no endpoint in endpoints.json at all, and the reproduction that found this bug
  // used --preset ...,calendar), and any preset added later starts out unlisted everywhere.
  it.each(FLAG_UNIVERSAL_NAMES)('%s is in every named preset once the flag is on', (tool) => {
    const missing = NAMED_PRESETS.filter((preset) => !presetContains(preset, tool, true));
    expect(missing, `${tool} missing from presets`).toEqual([]);
  });

  it('a flag-universal tool is a real utility tool, not a typo', () => {
    for (const tool of FLAG_UNIVERSAL_NAMES) {
      expect(UTILITY_NAMES, `${tool} is not a utility tool`).toContain(tool);
    }
  });

  // Which tools mint is a property of their implementations, so read it off them rather than
  // restating it: any utility tool whose execute body reaches the minting helper is one the flag
  // makes newly useful, and must therefore be declared flag-universal.
  //
  // This is also what keeps the two membership rules above from passing vacuously. They iterate
  // FLAG_UNIVERSAL_UTILITY_TOOLS, so emptying that declaration - the exact regression that would
  // restore the bug - would otherwise leave them as zero-case loops that report success.
  it('every utility tool that mints is declared flag-universal', () => {
    const minting = UTILITY_TOOLS.filter((utility) =>
      /mintDownloadUrl\s*\(/.test(utility.execute.toString())
    ).map((utility) => utility.name);

    expect(
      minting.length,
      'no utility tool appears to call the minting helper - this detection has gone stale and every ' +
        'flag-universal rule in this file is now vacuous'
    ).toBeGreaterThan(0);
    expect(
      minting.filter((name) => !FLAG_UNIVERSAL_NAMES.includes(name)),
      'tool mints under --enable-attachment-urls but is not in FLAG_UNIVERSAL_UTILITY_TOOLS'
    ).toEqual([]);
  });

  // The reproduction, exactly: the mail-focused preset set that was deployed with the flag.
  it('the mail-focused preset set carries get-download-url only with the flag', () => {
    const presets = ['mail', 'calendar', 'tasks', 'contacts'];
    expect(toolsMatching(getCombinedPresetPattern(presets, { attachmentUrls: true }))).toContain(
      'get-download-url'
    );
    expect(toolsMatching(getCombinedPresetPattern(presets))).not.toContain('get-download-url');
    // download-bytes was always there; the bug was never that the preset had no downloader.
    expect(toolsMatching(getCombinedPresetPattern(presets))).toContain('download-bytes');
  });
});

describe('the flag adds tools, and only tools it has to', () => {
  it('changes nothing in the presets that already carried the minting tool', () => {
    for (const preset of ['files', 'onedrive', 'personal', 'work', 'search']) {
      const off = toolsMatching(getCategoryPattern(preset, {})!).sort();
      const on = toolsMatching(getCategoryPattern(preset, { attachmentUrls: true })!).sort();
      expect(on, `${preset} tool set changed under the flag`).toEqual(off);
    }
  });

  it('adds nothing but the flag-universal tools to any preset', () => {
    for (const preset of NAMED_PRESETS) {
      const off = new Set(toolsMatching(getCategoryPattern(preset, {})!));
      const added = toolsMatching(getCategoryPattern(preset, { attachmentUrls: true })!).filter(
        (name) => !off.has(name)
      );
      expect(added.sort(), `${preset} gained unexpected tools`).toEqual(
        FLAG_UNIVERSAL_NAMES.filter((name) => !off.has(name)).sort()
      );
    }
  });

  it('removes nothing from any preset', () => {
    for (const preset of NAMED_PRESETS) {
      const on = new Set(toolsMatching(getCategoryPattern(preset, { attachmentUrls: true })!));
      const removed = toolsMatching(getCategoryPattern(preset, {})!).filter(
        (name) => !on.has(name)
      );
      expect(removed, `${preset} lost tools under the flag`).toEqual([]);
    }
  });

  // TOOL_CATEGORIES is the no-flag view used for descriptions and preset validation; it must stay
  // exactly what an unflagged run gets, or the two views drift and one of them starts lying.
  it('TOOL_CATEGORIES is the no-flag view', () => {
    for (const preset of NAMED_PRESETS) {
      expect(TOOL_CATEGORIES[preset].pattern.source).toBe(getCategoryPattern(preset, {})!.source);
    }
  });
});

describe('parse-teams-url scoping (rechecked, not assumed)', () => {
  // The other scoped utility. It converts a Teams meeting URL into the joinWebUrl that the
  // onlineMeetings endpoints take, so its reach is exactly the presets those endpoints live in.
  // Derived from endpoints.json for the same reason as above: if an onlineMeetings endpoint is ever
  // given another preset, this fails instead of quietly under-scoping the parser.
  const onlineMeetingPresets = [
    ...new Set(
      endpoints
        .filter((endpoint) => /\/onlineMeetings\//.test(endpoint.pathPattern))
        .flatMap((endpoint) => endpoint.presets ?? [])
    ),
  ].sort();

  it('covers every preset that has an onlineMeetings endpoint', () => {
    expect(onlineMeetingPresets.length).toBeGreaterThan(0);
    const missing = onlineMeetingPresets.filter(
      (preset) => !presetContains(preset, 'parse-teams-url', false)
    );
    expect(missing, 'parse-teams-url missing from onlineMeetings presets').toEqual([]);
  });

  it('is not widened by the attachment-URL flag', () => {
    // parse-teams-url mints nothing and reads no bytes, so no flag changes what it can reach.
    for (const preset of NAMED_PRESETS) {
      expect(
        presetContains(preset, 'parse-teams-url', true),
        `parse-teams-url membership of ${preset} changed under the flag`
      ).toBe(presetContains(preset, 'parse-teams-url', false));
    }
  });
});

describe('the registration selector is the same one the startup check asks', () => {
  let server: McpServer;
  let toolSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    vi.clearAllMocks();
    server = new McpServer({ name: 'test', version: '1.0.0' });
    toolSpy = vi.spyOn(server, 'tool').mockImplementation(() => {});
    vi.spyOn(server, 'registerTool').mockImplementation(
      () => ({}) as ReturnType<McpServer['registerTool']>
    );
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  function registeredUtilities(gates: {
    readOnly?: boolean;
    httpMode?: boolean;
    enabledTools?: string;
  }): string[] {
    registerGraphTools(
      server,
      {} as GraphClient,
      Boolean(gates.readOnly),
      gates.enabledTools,
      false,
      undefined,
      false,
      [],
      undefined,
      Boolean(gates.httpMode)
    );
    return toolSpy.mock.calls.map((call) => call[0] as string).sort();
  }

  const GATE_MATRIX: Array<{ name: string; gates: Record<string, unknown> }> = [
    { name: 'no filter', gates: {} },
    { name: 'read-only', gates: { readOnly: true } },
    { name: 'http mode', gates: { httpMode: true } },
    {
      name: 'mail-focused preset, flag off',
      gates: {
        httpMode: true,
        enabledTools: getCombinedPresetPattern(['mail', 'calendar', 'tasks', 'contacts']),
      },
    },
    {
      name: 'mail-focused preset, flag on',
      gates: {
        httpMode: true,
        enabledTools: getCombinedPresetPattern(['mail', 'calendar', 'tasks', 'contacts'], {
          attachmentUrls: true,
        }),
      },
    },
    { name: 'hand-written filter that drops the minting tool', gates: { enabledTools: '^send-' } },
    { name: 'filter matching nothing', gates: { enabledTools: '^no-such-tool$' } },
    { name: 'uncompilable filter', gates: { enabledTools: '[unclosed' } },
  ];

  // The startup warning is only as honest as this equality. If the predicate and the registration
  // loop can disagree, the server can warn about a tool it registered, or stay silent about one it
  // dropped -- and the second of those is the bug this whole file exists for.
  it.each(GATE_MATRIX)('$name: predicate matches what actually registers', ({ gates }) => {
    const actual = registeredUtilities(gates);
    expect(
      selectUtilityTools(gates)
        .map((utility) => utility.name)
        .sort()
    ).toEqual(actual);
    for (const name of UTILITY_NAMES) {
      expect(utilityToolWillRegister(name, gates), `${name} under ${JSON.stringify(gates)}`).toBe(
        actual.includes(name)
      );
    }
  });

  // The reproduction at the registration layer rather than the pattern layer.
  it('registers get-download-url under a mail-focused preset once the flag is on', () => {
    const registered = registeredUtilities({
      httpMode: true,
      enabledTools: getCombinedPresetPattern(['mail', 'calendar', 'tasks', 'contacts'], {
        attachmentUrls: true,
      }),
    });
    expect(registered).toContain('get-download-url');
    expect(registered).toContain('download-bytes');
  });
});

describe('startup warning when the flag has no tool to act through', () => {
  const savedEnv = { ...process.env };
  let started: MicrosoftGraphServer[] = [];

  beforeEach(() => {
    vi.clearAllMocks();
    process.env.MS365_MCP_RATE_LIMIT_DISABLED = 'true';
    process.env.MS365_MCP_ATTACHMENT_URL_KEY = 'shared-hmac-key';
  });

  afterEach(async () => {
    for (const server of started) await server.stop();
    started = [];
    resetAttachmentMinting();
    vi.restoreAllMocks();
    process.env = { ...savedEnv };
  });

  async function reserveFreePort(): Promise<number> {
    const holder = await new Promise<Server>((resolve) => {
      const s = createServer();
      s.listen(0, '127.0.0.1', () => resolve(s));
    });
    const port = (holder.address() as AddressInfo).port;
    await new Promise<void>((resolve) => holder.close(() => resolve()));
    return port;
  }

  async function startWith(enabledTools: string | undefined): Promise<void> {
    const port = await reserveFreePort();
    process.env.MS365_MCP_ATTACHMENT_URL_BASE = `http://127.0.0.1:${port}`;
    const options: CommandOptions = {
      http: `127.0.0.1:${port}`,
      trustProxyAuth: true,
      enableAuthTools: true,
      enableAttachmentUrls: true,
      enabledTools,
    };
    const server = new MicrosoftGraphServer(
      {
        isOAuthModeEnabled: () => false,
        isMultiAccount: async () => false,
        listAccounts: async () => [],
        getToken: async () => 'SERVER_OWN_TOKEN',
        getTokenForAccount: async () => 'SERVER_OWN_TOKEN',
        setOAuthToken: async () => {},
      } as unknown as AuthManager,
      options
    );
    await server.initialize('0.0.0-test');
    started.push(server);
    await server.start();
  }

  function unreachableWarnings(): string[] {
    return (logger.warn as unknown as ReturnType<typeof vi.fn>).mock.calls
      .map((call) => String(call[0]))
      .filter((message) => message.includes('get-download-url'));
  }

  it('warns in plain terms when a filter dropped the only tool that mints', async () => {
    await startWith('^(?:list-mail-messages|download-bytes)$');

    const warnings = unreachableWarnings();
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toContain('--enable-attachment-urls is set but get-download-url is NOT');
    expect(warnings[0]).toContain('cannot mint');
    // Name the knob that caused it, so the operator can act without reading the source.
    expect(warnings[0]).toContain('--enabled-tools');
  });

  it('stays silent when the mail-focused preset is built with the flag', async () => {
    await startWith(
      getCombinedPresetPattern(['mail', 'calendar', 'tasks', 'contacts'], { attachmentUrls: true })
    );

    expect(unreachableWarnings()).toEqual([]);
  });

  it('stays silent with no filter at all', async () => {
    await startWith(undefined);

    expect(unreachableWarnings()).toEqual([]);
  });
});
