/**
 * The startup check that warns when a flag has no tool left to act through asks
 * `utilityToolWillRegister`, and registration asks `selectUtilityTools`. The warning is only as
 * honest as the equality between those and what `registerGraphTools` actually registers: if
 * they can disagree, the server can warn about a tool it registered, or stay silent about one it
 * dropped. Which presets carry which flag-gated tool is covered in presets.test.ts and cli.test.ts.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import {
  registerGraphTools,
  selectUtilityTools,
  utilityToolWillRegister,
  UTILITY_TOOLS,
  type UtilityToolGates,
} from '../src/graph-tools.js';
import { getCombinedPresetPattern } from '../src/tool-categories.js';
import type GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn(), verbose: vi.fn() },
  enableConsoleLogging: vi.fn(),
}));

const MAIL_PRESETS = ['mail', 'calendar', 'tasks', 'contacts'];

const GATE_MATRIX: Array<{ name: string; gates: UtilityToolGates }> = [
  { name: 'no filter', gates: {} },
  { name: 'read-only', gates: { readOnly: true } },
  { name: 'http mode', gates: { httpMode: true } },
  {
    name: 'mail-focused preset, attachment URLs on',
    gates: {
      httpMode: true,
      enabledTools: getCombinedPresetPattern(MAIL_PRESETS, { attachmentUrls: true }),
    },
  },
  {
    name: 'mail-focused preset, proxy on',
    gates: {
      httpMode: true,
      attachmentProxy: true,
      enabledTools: getCombinedPresetPattern(MAIL_PRESETS, {
        attachmentUrls: true,
        attachmentProxy: true,
      }),
    },
  },
  { name: 'proxy on, no filter', gates: { httpMode: true, attachmentProxy: true } },
  { name: 'proxy on, read-only', gates: { readOnly: true, attachmentProxy: true } },
  {
    name: 'proxy on, filter that drops read-document',
    gates: { httpMode: true, attachmentProxy: true, enabledTools: '^send-' },
  },
  { name: 'filter matching nothing', gates: { enabledTools: '^no-such-tool$' } },
  { name: 'uncompilable filter', gates: { enabledTools: '[unclosed' } },
];

describe('the registration selector is the same one the startup check asks', () => {
  let server: McpServer;
  let toolSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    server = new McpServer({ name: 'test', version: '1.0.0' });
    toolSpy = vi.spyOn(server, 'tool').mockImplementation(() => {});
    vi.spyOn(server, 'registerTool').mockImplementation(
      () => ({}) as ReturnType<McpServer['registerTool']>
    );
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  const utilityNames = new Set(UTILITY_TOOLS.map((u) => u.name));

  it.each(GATE_MATRIX)('$name: predicate matches what actually registers', ({ gates }) => {
    registerGraphTools(server, {} as GraphClient, gates);
    const registeredUtilities = toolSpy.mock.calls
      .map((call) => call[0] as string)
      .filter((name) => utilityNames.has(name))
      .sort();

    expect(
      selectUtilityTools(gates)
        .map((utility) => utility.name)
        .sort()
    ).toEqual(registeredUtilities);
    for (const name of utilityNames) {
      expect(utilityToolWillRegister(name, gates), `${name} under ${JSON.stringify(gates)}`).toBe(
        registeredUtilities.includes(name)
      );
    }
  });

  it('registers read-document and not the byte readers once the proxy is on', () => {
    const gates = { httpMode: true, attachmentProxy: true };
    expect(utilityToolWillRegister('read-document', gates)).toBe(true);
    expect(utilityToolWillRegister('download-bytes', gates)).toBe(false);
    expect(utilityToolWillRegister('get-download-url', gates)).toBe(false);
  });

  it('keeps read-document out when the proxy is off', () => {
    expect(utilityToolWillRegister('read-document', { httpMode: true })).toBe(false);
  });
});
