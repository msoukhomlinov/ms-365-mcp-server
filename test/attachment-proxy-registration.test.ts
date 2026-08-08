/**
 * What --attachment-proxy registers, and what it takes away.
 *
 * Three tools go: download-bytes and get-download-url (utilities), and
 * get-mail-message-mime (a Graph endpoint). One arrives: read-document. Net -2.
 *
 * get-mail-message-mime is selected by a CLASS rule, not by its name. It is a
 * GET whose path ends in `/$value`, which is Graph's own spelling for "the raw
 * bytes of this resource", and any GET upstream adds with that shape is
 * suppressed the day it lands with no edit here. It also cannot be left to the
 * response scrubber: it is declared `acceptType: "text/plain"` and returns RFC
 * 5322 source, which is not itself base64 and so matches neither scrubber rule,
 * while carrying every attachment inline.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createServer, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import {
  buildToolsRegistry,
  isProxySuppressedGraphTool,
  registerGraphTools,
  selectUtilityTools,
  utilityToolWillRegister,
} from '../src/graph-tools.js';
import { getCombinedPresetPattern } from '../src/tool-categories.js';
import type GraphClient from '../src/graph-client.js';
import MicrosoftGraphServer from '../src/server.js';
import type AuthManager from '../src/auth.js';
import type { CommandOptions } from '../src/cli.js';
import { resetAttachmentMinting } from '../src/lib/attachment-minting.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn(), verbose: vi.fn() },
  enableConsoleLogging: vi.fn(),
}));

import logger from '../src/logger.js';

const BYTE_TOOLS = ['download-bytes', 'get-download-url', 'get-mail-message-mime'];

describe('--attachment-proxy registration surface', () => {
  let server: McpServer;
  let utilitySpy: ReturnType<typeof vi.spyOn>;
  let graphSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    vi.clearAllMocks();
    server = new McpServer({ name: 'test', version: '1.0.0' });
    utilitySpy = vi.spyOn(server, 'tool').mockImplementation(() => ({}) as never);
    graphSpy = vi.spyOn(server, 'registerTool').mockImplementation(() => ({}) as never);
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  /** Every tool name this configuration registers, utilities and Graph alike. */
  function register(attachmentProxy: boolean): { names: string[]; count: number } {
    const count = registerGraphTools(
      server,
      {} as GraphClient,
      false, // readOnly
      undefined, // enabledToolsPattern
      false, // orgMode
      undefined, // authManager
      false, // multiAccount
      [], // accountNames
      undefined, // allowedScopes
      true, // httpMode -- proxy mode is HTTP only
      attachmentProxy
    );
    const names = [
      ...utilitySpy.mock.calls.map((call) => call[0] as string),
      ...graphSpy.mock.calls.map((call) => call[0] as string),
    ];
    return { names, count };
  }

  it('registers read-document and none of the three byte tools', () => {
    const { names } = register(true);
    expect(names).toContain('read-document');
    for (const tool of BYTE_TOOLS) {
      expect(names, `${tool} must not register under --attachment-proxy`).not.toContain(tool);
    }
  });

  it('leaves all four exactly as upstream when the flag is off', () => {
    const { names } = register(false);
    expect(names).not.toContain('read-document');
    for (const tool of BYTE_TOOLS) {
      expect(names, `${tool} must register without the flag`).toContain(tool);
    }
  });

  it('removes exactly two tools net', () => {
    // Asserted as a delta, not as "117". The absolute number depends on the
    // preset and on how many endpoints ship in a given release, so a literal
    // would be a maintenance tax that fails on an unrelated upstream merge.
    // The deployed 119 -> 117 is verified against the built image at rollout.
    const off = register(false).count;
    vi.clearAllMocks();
    server = new McpServer({ name: 'test', version: '1.0.0' });
    utilitySpy = vi.spyOn(server, 'tool').mockImplementation(() => ({}) as never);
    graphSpy = vi.spyOn(server, 'registerTool').mockImplementation(() => ({}) as never);
    const on = register(true).count;
    expect(on).toBe(off - 2);
  });

  it('keeps the predicate and the registration loop in agreement', () => {
    // The startup warning asks selectUtilityTools; if it can disagree with what
    // registers, the server can warn about a tool it registered or stay silent
    // about one it dropped.
    const gates = { httpMode: true, attachmentProxy: true };
    const { names } = register(true);
    const utilities = selectUtilityTools(gates).map((u) => u.name);
    for (const name of [...utilities, ...BYTE_TOOLS, 'read-document']) {
      if (name === 'get-mail-message-mime') continue; // Graph endpoint, not a utility
      expect(utilityToolWillRegister(name, gates), `${name}`).toBe(names.includes(name));
    }
  });
});

describe('the $value suppression is a class rule', () => {
  it('selects any GET whose path ends in /$value, and nothing else', () => {
    expect(isProxySuppressedGraphTool('get', '/me/messages/{message-id}/$value')).toBe(true);
    expect(isProxySuppressedGraphTool('GET', '/me/photo/$value')).toBe(true);
    // An upload to the same path is not a byte READ.
    expect(isProxySuppressedGraphTool('put', '/me/photo/$value')).toBe(false);
    expect(isProxySuppressedGraphTool('get', '/me/messages')).toBe(false);
    expect(isProxySuppressedGraphTool('get', '/me/messages/{id}/$value/extra')).toBe(false);
    expect(isProxySuppressedGraphTool('get', undefined)).toBe(false);
  });
});

describe('discovery mode suppresses the same reads', () => {
  it('drops get-mail-message-mime from the discovery registry under the flag', () => {
    const off = buildToolsRegistry(false, false, undefined, undefined, [], false);
    const on = buildToolsRegistry(false, false, undefined, undefined, [], true);
    expect(off.has('get-mail-message-mime')).toBe(true);
    expect(on.has('get-mail-message-mime')).toBe(false);
    // Nothing else changed: search-tools would otherwise advertise a surface
    // that execute-tool cannot run, or hide one it can.
    expect(on.size).toBe(off.size - 1);
  });
});

describe('startup warnings around --attachment-proxy', () => {
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

  function fakeAuthManager(): AuthManager {
    return {
      isOAuthModeEnabled: () => false,
      isMultiAccount: async () => false,
      listAccounts: async () => [],
      getToken: async () => 'SERVER_OWN_TOKEN',
      getTokenForAccount: async () => 'SERVER_OWN_TOKEN',
      setOAuthToken: async () => {},
    } as unknown as AuthManager;
  }

  async function reserveFreePort(): Promise<number> {
    const holder = await new Promise<Server>((resolve) => {
      const s = createServer();
      s.listen(0, '127.0.0.1', () => resolve(s));
    });
    const port = (holder.address() as AddressInfo).port;
    await new Promise<void>((resolve) => holder.close(() => resolve()));
    return port;
  }

  async function start(options: CommandOptions): Promise<void> {
    const server = new MicrosoftGraphServer(fakeAuthManager(), options);
    await server.initialize('0.0.0-test');
    started.push(server);
    await server.start();
  }

  function warnings(): string[] {
    return (logger.warn as unknown as ReturnType<typeof vi.fn>).mock.calls.map((call) =>
      String(call[0])
    );
  }

  it('warns and does nothing else when --attachment-proxy is set in stdio mode', async () => {
    // No --http at all: registerGraphTools never even sees attachmentProxy=true (see
    // createMcpServer's attachmentProxyActive gate), so every byte tool a plain stdio run
    // would register stays registered. The warning is the only observable effect.
    await start({ attachmentProxy: 'http://docglean:8080/mcp', enableAttachmentUrls: true });

    const w = warnings();
    expect(w.some((m) => /--attachment-proxy has no effect in stdio mode/.test(m))).toBe(true);
    // The flag it implies keeps its own warning too; neither swallows the other.
    expect(w.some((m) => /--enable-attachment-urls has no effect in stdio mode/.test(m))).toBe(
      true
    );
  });

  it('does not warn that get-download-url is unreachable when the proxy deliberately removed it', async () => {
    const port = await reserveFreePort();
    process.env.MS365_MCP_ATTACHMENT_URL_BASE = `http://127.0.0.1:${port}`;
    await start({
      http: `127.0.0.1:${port}`,
      trustProxyAuth: true,
      attachmentProxy: 'http://docglean:8080/mcp',
      enableAttachmentUrls: true,
      enabledTools: getCombinedPresetPattern(['mail', 'calendar', 'tasks', 'contacts'], {
        attachmentProxy: true,
      }),
    });

    // Without the attachmentProxyActive exemption this would fire: attachmentProxy implies
    // enableAttachmentUrls, and get-download-url is deliberately not registered under the proxy.
    expect(warnings().some((m) => /get-download-url is NOT registered/.test(m))).toBe(false);
  });

  it('warns when the active filter drops read-document, leaving no readable tool at all', async () => {
    const port = await reserveFreePort();
    process.env.MS365_MCP_ATTACHMENT_URL_BASE = `http://127.0.0.1:${port}`;
    await start({
      http: `127.0.0.1:${port}`,
      trustProxyAuth: true,
      attachmentProxy: 'http://docglean:8080/mcp',
      enableAttachmentUrls: true,
      enabledTools: '^(?:list-mail-messages)$',
    });

    const w = warnings().filter((m) => m.includes('read-document'));
    expect(w).toHaveLength(1);
    expect(w[0]).toContain('--attachment-proxy is set but read-document is NOT registered');
    expect(w[0]).toContain('--enabled-tools');
  });

  it('stays silent about read-document when a preset built with the flag carries it', async () => {
    const port = await reserveFreePort();
    process.env.MS365_MCP_ATTACHMENT_URL_BASE = `http://127.0.0.1:${port}`;
    await start({
      http: `127.0.0.1:${port}`,
      trustProxyAuth: true,
      attachmentProxy: 'http://docglean:8080/mcp',
      enableAttachmentUrls: true,
      enabledTools: getCombinedPresetPattern(['mail', 'calendar', 'tasks', 'contacts'], {
        attachmentProxy: true,
      }),
    });

    expect(warnings().some((m) => m.includes('read-document is NOT registered'))).toBe(false);
  });
});
