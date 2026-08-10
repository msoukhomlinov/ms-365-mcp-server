/**
 * Guidance text may never name a tool this configuration did not register.
 *
 * `--attachment-proxy` deletes three tools (download-bytes, get-download-url,
 * get-mail-message-mime) and adds read-document. The descriptions, parameter
 * descriptions and `initialize.instructions` are built from the same static
 * strings in every mode, so every sentence telling the model to "call
 * download-bytes" survives into a server that has no such tool. A model that
 * follows the guidance calls a name that is not there.
 *
 * These tests assert the PROPERTY, not one string: collect every piece of text
 * this server hands the model, extract every known tool name it mentions, and
 * require each mentioned name to be registered. A hardcoded-string assertion
 * would pass the day someone adds the next llmTip.
 */
import { describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import {
  buildToolsRegistry,
  proxySuppressedToolNames,
  registerDiscoveryTools,
  registerGraphTools,
  resolveRegisteredToolNames,
  selectUtilityTools,
  UTILITY_TOOLS,
  type UtilityToolGates,
} from '../src/graph-tools.js';
import { buildMcpServerInstructions } from '../src/mcp-instructions.js';
import { mentionsToolName, stripStaleToolGuidance } from '../src/lib/guidance-text.js';
import type GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn(), verbose: vi.fn() },
  enableConsoleLogging: vi.fn(),
}));

const endpointsData = JSON.parse(
  readFileSync(path.join(import.meta.dirname, '..', 'src', 'endpoints.json'), 'utf8')
) as Array<{ toolName: string }>;

/** Registered by registerAuthTools in every mode, so never a stale reference. */
const AUTH_TOOL_NAMES = [
  'login',
  'logout',
  'verify-login',
  'list-accounts',
  'select-account',
  'remove-account',
];

/** Registered only by registerDiscoveryTools. */
const DISCOVERY_TOOL_NAMES = ['search-tools', 'get-tool-schema', 'execute-tool'];

/**
 * Every tool name any build of this server can register. Single-word names
 * (login, logout) are excluded: they are ordinary English and would match
 * prose, and they are registered unconditionally anyway.
 */
const KNOWN_TOOL_NAMES: string[] = [
  ...endpointsData.map((e) => e.toolName),
  ...UTILITY_TOOLS.map((u) => u.name),
  ...AUTH_TOOL_NAMES,
  ...DISCOVERY_TOOL_NAMES,
].filter((name) => name.includes('-'));

/**
 * Tool names mentioned in `text`. Matched on word boundaries so
 * "download-bytes" does not match inside "download-bytes-to-file", and
 * separator-agnostically because MCP clients render these names with
 * underscores (the live server's list-mail-attachments tip reads
 * "download_bytes" in the client).
 */
function mentionedToolNames(text: string): string[] {
  const found = new Set<string>();
  for (const name of KNOWN_TOOL_NAMES) {
    const pattern = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/-/g, '[-_]');
    if (new RegExp(`(?<![A-Za-z0-9_-])${pattern}(?![A-Za-z0-9_-])`, 'i').test(text)) {
      found.add(name);
    }
  }
  return [...found];
}

type Surface = { tool: string; where: string; text: string };

/** Parameter descriptions, for both the raw-shape and z.object() registration forms. */
function paramDescriptions(schema: unknown): Array<{ param: string; text: string }> {
  const out: Array<{ param: string; text: string }> = [];
  if (!schema || typeof schema !== 'object') return out;
  const def = (schema as { _def?: { shape?: unknown } })._def;
  let shape: Record<string, unknown> | undefined;
  if (def && typeof def.shape === 'function') {
    shape = (def.shape as () => Record<string, unknown>)();
  } else if (def && def.shape && typeof def.shape === 'object') {
    shape = def.shape as Record<string, unknown>;
  } else {
    shape = schema as Record<string, unknown>;
  }
  for (const [param, zodType] of Object.entries(shape ?? {})) {
    const description = (zodType as { description?: unknown })?.description;
    if (typeof description === 'string' && description.length > 0) {
      out.push({ param, text: description });
    }
  }
  return out;
}

/**
 * Registers for real and captures what reached the MCP SDK. Only the SDK
 * boundary is stubbed — the descriptions, llmTip assembly and every gate run
 * as they do in production.
 */
type Handler = (params: Record<string, unknown>) => Promise<{ content: Array<{ text: string }> }>;

function registerAndCollect(opts: {
  discovery: boolean;
  attachmentProxy: boolean;
  httpMode: boolean;
  enabledTools?: string;
}): { names: Set<string>; surfaces: Surface[]; handlers: Map<string, Handler> } {
  const server = new McpServer({ name: 'test', version: '1.0.0' });
  const surfaces: Surface[] = [];
  const names = new Set<string>([...AUTH_TOOL_NAMES]);
  const handlers = new Map<string, Handler>();

  const record = (name: string, description: string, schema: unknown, args: unknown[]) => {
    names.add(name);
    if (typeof description === 'string') {
      surfaces.push({ tool: name, where: 'description', text: description });
    }
    for (const { param, text } of paramDescriptions(schema)) {
      surfaces.push({ tool: name, where: `param:${param}`, text });
    }
    const last = args[args.length - 1];
    if (typeof last === 'function') handlers.set(name, last as Handler);
  };

  vi.spyOn(server, 'tool').mockImplementation(((...args: unknown[]) => {
    record(args[0] as string, args[1] as string, args[2], args);
    return {} as never;
  }) as never);
  vi.spyOn(server, 'registerTool').mockImplementation(((...args: unknown[]) => {
    const config = args[1] as { description?: string; inputSchema?: unknown };
    record(args[0] as string, config?.description ?? '', config?.inputSchema, args);
    return {} as never;
  }) as never);

  if (opts.discovery) {
    // Discovery mode registers three meta-tools; the Graph tools and utilities
    // it can reach are called through execute-tool rather than registered
    // individually. They count as present for this property — naming one is not
    // a stale reference — while the tools the flag removed are in neither set.
    for (const name of buildToolsRegistry(
      false,
      true,
      opts.enabledTools ? new RegExp(opts.enabledTools, 'i') : undefined,
      undefined,
      [],
      opts.attachmentProxy
    ).keys()) {
      names.add(name);
    }
    for (const utility of selectUtilityTools({
      httpMode: opts.httpMode,
      attachmentProxy: opts.attachmentProxy,
      enabledTools: opts.enabledTools,
    })) {
      names.add(utility.name);
    }
    registerDiscoveryTools(
      server,
      {} as GraphClient,
      false, // readOnly
      true, // orgMode
      undefined, // authManager
      false, // multiAccount
      [], // accountNames
      opts.enabledTools,
      undefined, // allowedScopes
      opts.httpMode,
      opts.attachmentProxy, // attachmentUrls -- proxy implies it
      opts.attachmentProxy
    );
  } else {
    registerGraphTools(
      server,
      {} as GraphClient,
      false, // readOnly
      opts.enabledTools,
      true, // orgMode
      undefined, // authManager
      false, // multiAccount
      [], // accountNames
      undefined, // allowedScopes
      opts.httpMode,
      opts.attachmentProxy
    );
  }

  vi.restoreAllMocks();
  return { names, surfaces, handlers };
}

/** Every mentioned-but-unregistered reference, as readable lines. */
function staleReferences(names: Set<string>, surfaces: Surface[]): string[] {
  const stale: string[] = [];
  for (const surface of surfaces) {
    for (const mentioned of mentionedToolNames(surface.text)) {
      if (!names.has(mentioned)) {
        stale.push(`${surface.tool} (${surface.where}) -> ${mentioned}`);
      }
    }
  }
  return [...new Set(stale)].sort();
}

const PROXY_GATES: UtilityToolGates = { httpMode: true, attachmentProxy: true };

describe('guidance text under --attachment-proxy', () => {
  it('names no unregistered tool in any registered tool description or parameter', () => {
    const { names, surfaces } = registerAndCollect({
      discovery: false,
      attachmentProxy: true,
      httpMode: true,
    });

    expect(names.has('read-document')).toBe(true);
    expect(names.has('download-bytes')).toBe(false);
    expect(surfaces.length).toBeGreaterThan(100);
    expect(staleReferences(names, surfaces)).toEqual([]);
  });

  it('names no unregistered tool in discovery mode either', () => {
    const { names, surfaces } = registerAndCollect({
      discovery: true,
      attachmentProxy: true,
      httpMode: true,
    });

    expect(names.has('search-tools')).toBe(true);
    expect(staleReferences(names, surfaces)).toEqual([]);
  });

  it('names no unregistered tool in the MCP instructions', () => {
    const { names } = registerAndCollect({
      discovery: false,
      attachmentProxy: true,
      httpMode: true,
    });
    const instructions = buildMcpServerInstructions({
      orgMode: true,
      readOnly: false,
      multiAccount: false,
      discovery: false,
      registeredTools: resolveRegisteredToolNames({
        orgMode: true,
        httpMode: true,
        attachmentProxy: true,
      }),
    });

    expect(
      staleReferences(names, [{ tool: 'instructions', where: 'text', text: instructions }])
    ).toEqual([]);
    expect(instructions).toContain('read-document');
  });

  it('still routes byte reads to the byte tools when no proxy is configured', () => {
    const instructions = buildMcpServerInstructions({
      orgMode: true,
      readOnly: false,
      multiAccount: false,
      discovery: false,
      registeredTools: resolveRegisteredToolNames({ orgMode: true, httpMode: false }),
    });

    expect(instructions).toContain('get-download-url');
    expect(instructions).toContain('download-bytes');
    // stdio-only guidance is true in stdio mode and must not be scrubbed away.
    expect(instructions).toContain('download-bytes-to-file');
    expect(instructions).not.toContain('read-document');
  });

  it('keeps byte-tool guidance intact in a non-proxy Graph registration', () => {
    const { names, surfaces } = registerAndCollect({
      discovery: false,
      attachmentProxy: false,
      httpMode: false,
    });

    expect(names.has('download-bytes')).toBe(true);
    const attachmentsTip = surfaces.find(
      (s) => s.tool === 'list-mail-attachments' && s.where === 'description'
    );
    expect(attachmentsTip?.text).toContain('download-bytes');
    expect(staleReferences(names, surfaces)).toEqual([]);
  });

  // In discovery mode the llmTips reach the model through these two handlers,
  // not through a registered description, so the property has to be checked on
  // what they return. Mocking the registry instead of the SDK boundary here
  // would exercise none of the real assembly.
  it('names no unregistered tool in search-tools or get-tool-schema results', async () => {
    const { names, handlers } = registerAndCollect({
      discovery: true,
      attachmentProxy: true,
      httpMode: true,
    });

    const search = await handlers.get('search-tools')!({
      query: 'download mail attachment bytes photo hosted content drive file',
      limit: 50,
    });
    const attachments = await handlers.get('get-tool-schema')!({
      tool_name: 'list-mail-attachments',
    });
    const driveItem = await handlers.get('get-tool-schema')!({ tool_name: 'get-drive-item' });
    const readDocument = await handlers.get('get-tool-schema')!({ tool_name: 'read-document' });

    const surfaces: Surface[] = [
      { tool: 'search-tools', where: 'result', text: search.content[0].text },
      {
        tool: 'get-tool-schema',
        where: 'list-mail-attachments',
        text: attachments.content[0].text,
      },
      { tool: 'get-tool-schema', where: 'get-drive-item', text: driveItem.content[0].text },
      { tool: 'get-tool-schema', where: 'read-document', text: readDocument.content[0].text },
    ];
    expect(surfaces.every((s) => s.text.length > 0)).toBe(true);
    expect(staleReferences(names, surfaces)).toEqual([]);
  });

  it('derives the suppressed set from the tool table rather than a literal list', () => {
    const suppressed = proxySuppressedToolNames(PROXY_GATES);

    // Every utility the flag drops, by its bytesToModel marking rather than by name.
    for (const utility of UTILITY_TOOLS.filter((u) => u.bytesToModel)) {
      expect(suppressed).toContain(utility.name);
    }
    // And the Graph tool it drops by the GET-/$value class rule.
    expect(suppressed).toContain('get-mail-message-mime');
    // Not a proxy-mode casualty: stdio-only, and its guidance says so.
    expect(suppressed).not.toContain('download-bytes-to-file');
    // Nothing is suppressed when the flag is off, so no other mode pays for this.
    expect(proxySuppressedToolNames({ httpMode: true })).toEqual([]);
  });

  it('keeps the surviving guidance and states the replacement once', () => {
    const { surfaces } = registerAndCollect({
      discovery: false,
      attachmentProxy: true,
      httpMode: true,
    });
    const tip = surfaces.find(
      (s) => s.tool === 'list-mail-attachments' && s.where === 'description'
    )!.text;

    // The non-tool half of the tip survives...
    expect(tip).toContain('id, name, contentType, size, isInline');
    // ...the $value requirement the dropped sentence carried is restored...
    expect(tip).toContain('/$value');
    expect(tip).toContain('read-document');
    // ...and the replacement is stated once, not once per dropped sentence.
    expect(tip.match(/Byte payloads are stripped from every tool result/g)).toHaveLength(1);
  });
});

/**
 * PR #19 review, both findings. They are one defect: the suppressed set was a
 * delta between two `selectUtilityTools` calls and the instructions keyed off
 * proxy activation, when the only thing true is the resolved registered set.
 *
 * A tool excluded by proxy mode AND an --enabled-tools regex is missing from
 * both sides of that delta, so it cancels out and its stale guidance survives;
 * a tool the regex dropped is still advertised by the instructions. Both
 * expressions below are the reviewer's own.
 */
describe('guidance under --attachment-proxy combined with an --enabled-tools filter', () => {
  /** The reachable tool names for a configuration, from the gates registration uses. */
  function reachable(opts: {
    enabledTools?: string;
    attachmentProxy: boolean;
    httpMode: boolean;
  }): Set<string> {
    const names = new Set<string>(AUTH_TOOL_NAMES);
    for (const name of buildToolsRegistry(
      false,
      true,
      opts.enabledTools ? new RegExp(opts.enabledTools, 'i') : undefined,
      undefined,
      [],
      opts.attachmentProxy
    ).keys()) {
      names.add(name);
    }
    for (const utility of selectUtilityTools({
      httpMode: opts.httpMode,
      attachmentProxy: opts.attachmentProxy,
      enabledTools: opts.enabledTools,
    })) {
      names.add(utility.name);
    }
    return names;
  }

  // Finding 1. download-bytes is dropped twice over -- by proxy mode and by the
  // regex -- so a delta between the two selections cannot see it.
  it('drops byte-tool guidance when the filter also excludes the byte tool', () => {
    const enabledTools = '^(list-mail-attachments|read-document)$';
    const { names, surfaces } = registerAndCollect({
      discovery: false,
      attachmentProxy: true,
      httpMode: true,
      enabledTools,
    });

    expect(names.has('read-document')).toBe(true);
    expect(names.has('download-bytes')).toBe(false);
    expect(staleReferences(names, surfaces)).toEqual([]);
  });

  // Finding 2. The filter admits a Graph tool but not read-document: startup
  // warns and keeps running, so the instructions must not promise a tool that
  // was never registered.
  it('promises no document-read tool when the filter excludes read-document', () => {
    const enabledTools = '^(list-mail-attachments)$';
    const names = reachable({ enabledTools, attachmentProxy: true, httpMode: true });
    expect(names.has('read-document')).toBe(false);

    const instructions = buildMcpServerInstructions({
      orgMode: true,
      readOnly: false,
      multiAccount: false,
      discovery: false,
      registeredTools: names,
    });

    expect(
      staleReferences(names, [{ tool: 'instructions', where: 'text', text: instructions }])
    ).toEqual([]);
    expect(instructions).not.toContain('read-document');
  });

  // Same filter, the description surface rather than the instructions: the
  // replacement sentence names read-document, so it cannot be appended when
  // read-document is the tool that went missing.
  it('does not substitute read-document into descriptions when it is filtered out too', () => {
    const enabledTools = '^(list-mail-attachments)$';
    const { names, surfaces } = registerAndCollect({
      discovery: false,
      attachmentProxy: true,
      httpMode: true,
      enabledTools,
    });

    expect(names.has('read-document')).toBe(false);
    expect(staleReferences(names, surfaces)).toEqual([]);
  });

  /*
   * Why filtering stops at the proxy byte tools instead of every unregistered
   * name (PR #19 re-review, finding A).
   *
   * A sentence carries more than one fact. `update-planner-bucket`'s whole tip
   * is one sentence -- "CRITICAL: Requires If-Match header with ETag from
   * get-planner-bucket (use includeHeaders=true)." -- so dropping it to avoid
   * naming an unregistered get-planner-bucket takes the If-Match requirement
   * with it. Fourteen tips in endpoints.json put a requirement in the same
   * sentence as a cross-reference and two of them are single-sentence tips
   * where nothing survives the drop.
   *
   * The two failure modes are not equal. Naming a tool that is not registered
   * costs one failed tool call, with an error that says exactly what is wrong.
   * Losing an If-Match requirement produces a well-formed request that Graph
   * rejects with 412, or a todoTaskListId omitted from a create -- silent,
   * and blamed on Graph rather than on this text. Trading the legible failure
   * for the silent one is not a win, so the byte-tool guidance the proxy flag
   * strands is filtered and other cross-references are left alone.
   *
   * Presets are why the residual exposure is small: download-bytes and
   * download-bytes-to-file are in every preset (UNIVERSAL_UTILITY_TOOLS), so
   * only a hand-written --enabled-tools regex can strand a cross-referenced
   * tool. Splitting cross-reference from requirement in endpoints.json is the
   * fix that would make the wider invariant safe; it is not this change.
   */
  it('keeps a requirement whose sentence also names an excluded tool', () => {
    const { names, surfaces } = registerAndCollect({
      discovery: false,
      attachmentProxy: false,
      httpMode: false,
      enabledTools: '^update-planner-bucket$',
    });

    expect(names.has('update-planner-bucket')).toBe(true);
    expect(names.has('get-planner-bucket')).toBe(false);
    const tip = surfaces.find(
      (s) => s.tool === 'update-planner-bucket' && s.where === 'description'
    )!.text;

    // The requirement survives, cross-reference and all: losing it means a 412
    // the model cannot diagnose.
    expect(tip).toContain('If-Match');
    expect(tip).toContain('ETag');
    expect(tip).toContain('includeHeaders');
  });

  it('keeps the same requirement on delete-planner-bucket', () => {
    const { surfaces } = registerAndCollect({
      discovery: false,
      attachmentProxy: false,
      httpMode: false,
      enabledTools: '^delete-planner-bucket$',
    });
    const tip = surfaces.find(
      (s) => s.tool === 'delete-planner-bucket' && s.where === 'description'
    )!.text;

    expect(tip).toContain('If-Match');
    expect(tip).toContain('ETag');
  });

  // Byte-tool guidance is still filtered even when the same filter strands the
  // byte tool -- that is the defect this PR exists for, and it does not lose a
  // requirement, because a "call download-bytes with target=..." sentence is a
  // direction and nothing else.
  it('still drops byte-tool guidance the proxy flag stranded, under a filter', () => {
    const { names, surfaces } = registerAndCollect({
      discovery: false,
      attachmentProxy: true,
      httpMode: true,
      enabledTools: '^(list-mail-attachments|read-document)$',
    });

    expect(names.has('download-bytes')).toBe(false);
    const tip = surfaces.find(
      (s) => s.tool === 'list-mail-attachments' && s.where === 'description'
    )!.text;
    expect(tip).not.toMatch(/download[-_]bytes/i);
    expect(tip).toContain('read-document');
    // The replacement carries the same absence claim as the instructions, so it
    // is held to the same standard: raw bytes yes (scrubber), download URL no.
    expect(tip).not.toMatch(/or a download URL/i);
    expect(tip).toMatch(/raw bytes/i);
  });

  /*
   * The no-byte-tool paragraph claims only what those four utilities do (PR #19
   * re-review, finding B). Their absence says nothing about Graph tools that
   * return content inside their JSON: get-mail-message returns the message body
   * whether or not a byte tool exists, so "no content can be read" is false and
   * "do not offer a read" is an instruction to refuse work the server can do.
   *
   * Narrowed rather than enumerated. A list of content-returning Graph tools
   * would go stale against endpoints.json every time upstream adds one, and
   * being wrong in the other direction -- claiming a read is available when it
   * is not -- is the failure this whole PR is about.
   */
  /*
   * The fallback asserted an absence it could not verify (PR #19 re-review,
   * finding C -- the second catch of the same shape). With
   * `^get-meeting-recording-content$` the four utilities are all absent, but
   * that Graph tool is registered, its path ends `/content` rather than
   * `/$value` so nothing suppresses it, and graph-client.ts base64-encodes the
   * MP4 into contentBytes. The paragraph denied a binary read the server
   * performs.
   *
   * Enumerating registered Graph byte endpoints instead would go stale on every
   * endpoints.json addition, silently. So the rule is: guidance may state what
   * is registered, never what cannot be done -- and with nothing to state, it
   * says nothing.
   */
  it('says nothing about byte content when no byte/document utility is registered', () => {
    const registeredTools = resolveRegisteredToolNames({
      orgMode: true,
      enabledTools: '^get-meeting-recording-content$',
    });
    expect(registeredTools.has('get-meeting-recording-content')).toBe(true);
    for (const tool of ['read-document', 'download-bytes', 'get-download-url']) {
      expect(registeredTools.has(tool)).toBe(false);
    }

    const instructions = buildMcpServerInstructions({
      orgMode: true,
      readOnly: false,
      multiAccount: false,
      discovery: false,
      registeredTools,
    });

    // No byte-content guidance at all, rather than a claim of absence.
    expect(instructions).not.toContain('Files / binary content');
    expect(instructions).not.toMatch(/cannot be read/i);
    expect(instructions).not.toMatch(/no tool that returns/i);
    // Dropping a clause must not leave a seam behind.
    expect(instructions).not.toMatch(/ {2}/);
    expect(instructions).not.toMatch(/\.\s*\./);
    // The rest of the instructions are unaffected.
    expect(instructions).toContain('Microsoft Graph OData');
  });

  /*
   * The one absence claim that survives is the one a mechanism enforces: under
   * proxy mode installResponseScrubbing strips contentBytes and any base64 over
   * 4 KB from every tool result (src/lib/response-scrubber.ts), and it is
   * installed on exactly the condition that registers read-document. Download
   * URLs are NOT stripped, and get-drive-item still returns
   * @microsoft.graph.downloadUrl, so claiming none is available was false.
   */
  it('claims no raw bytes under proxy mode but does not deny download URLs', () => {
    const registeredTools = resolveRegisteredToolNames({
      orgMode: true,
      httpMode: true,
      attachmentProxy: true,
    });
    expect(registeredTools.has('read-document')).toBe(true);
    expect(registeredTools.has('get-drive-item')).toBe(true);

    const instructions = buildMcpServerInstructions({
      orgMode: true,
      readOnly: false,
      multiAccount: false,
      discovery: false,
      registeredTools,
    });

    // Scrubber-backed, so it may be stated.
    expect(instructions).toMatch(/raw bytes/i);
    // Not scrubber-backed: get-drive-item hands the model a download URL.
    expect(instructions).not.toMatch(/or a download URL/i);
    expect(instructions).not.toMatch(/mints a download URL/i);
  });

  /*
   * Finding D. read-document mints internally, so it inherits the identity
   * constraint get-download-url already carries (src/graph-tools.ts:704), and
   * its own guard at src/graph-tools.ts:1618 --
   * `ctx.authManager?.isOAuthModeEnabled() || getRequestTokens()` -- answers
   * identity_not_supported before any conversion. Under --http
   * --attachment-proxy with bearer, OAuth or OBO the tool is registered and
   * visible in the model's tool list but always refuses, so the guidance has to
   * name the condition.
   */
  it('qualifies read-document with the request-scoped identity condition', () => {
    const registeredTools = resolveRegisteredToolNames({
      orgMode: true,
      httpMode: true,
      attachmentProxy: true,
    });
    const instructions = buildMcpServerInstructions({
      orgMode: true,
      readOnly: false,
      multiAccount: false,
      discovery: false,
      registeredTools,
    });

    expect(instructions).toContain('read-document');
    expect(instructions).toMatch(/OAuth, OBO, or bearer mode/);
    expect(instructions).toMatch(/Authorization header/i);
  });

  it('claims only that byte/document tools are missing, not that content is unreadable', () => {
    const registeredTools = resolveRegisteredToolNames({
      orgMode: true,
      enabledTools: '^get-mail-message$',
    });
    expect(registeredTools.has('get-mail-message')).toBe(true);
    for (const tool of ['read-document', 'download-bytes', 'get-download-url']) {
      expect(registeredTools.has(tool)).toBe(false);
    }

    const instructions = buildMcpServerInstructions({
      orgMode: true,
      readOnly: false,
      multiAccount: false,
      discovery: false,
      registeredTools,
    });

    // Must not deny a capability the registered Graph tool plainly has.
    // get-mail-message returns the message body regardless of byte tools, and
    // narrowing the claim was not enough (see finding C): the paragraph is gone.
    expect(instructions).not.toContain('no document can be read here at all');
    expect(instructions).not.toContain('report the capability as unavailable');
    expect(instructions).not.toContain('Files / binary content');
    expect(instructions).not.toMatch(/cannot be read/i);
    expect(
      staleReferences(registeredTools, [{ tool: 'i', where: 't', text: instructions }])
    ).toEqual([]);
  });

  // The stdio-scoped download-bytes-to-file sentence: kept where the tool is
  // real, dropped where it is not. Registration decides, not the prose.
  it('keeps the stdio byte-to-file sentence in stdio and drops it over HTTP', () => {
    const stdio = reachable({ attachmentProxy: false, httpMode: false });
    const http = reachable({ attachmentProxy: false, httpMode: true });
    expect(stdio.has('download-bytes-to-file')).toBe(true);
    expect(http.has('download-bytes-to-file')).toBe(false);

    const base = { orgMode: true, readOnly: false, multiAccount: false, discovery: false };
    const stdioText = buildMcpServerInstructions({ ...base, registeredTools: stdio });
    const httpText = buildMcpServerInstructions({ ...base, registeredTools: http });

    expect(stdioText).toContain('download-bytes-to-file');
    expect(httpText).not.toContain('download-bytes-to-file');
    expect(httpText).toContain('download-bytes');
    // Which clause lands first depends on what registered, so the join has to
    // capitalise the ones that follow a full stop -- without renaming a tool
    // whose name happens to open the sentence.
    for (const text of [stdioText, httpText]) {
      expect(text).not.toMatch(/\.\s+[a-z]/);
      expect(text).not.toMatch(/[A-Z][a-z0-9]*-[a-z]+-/);
    }
    expect(staleReferences(http, [{ tool: 'i', where: 't', text: httpText }])).toEqual([]);
    expect(staleReferences(stdio, [{ tool: 'i', where: 't', text: stdioText }])).toEqual([]);
  });

  // The derivation itself: one definition, checked against what registration
  // actually did rather than against a second copy of the same reasoning.
  it('resolves the same tool names registration registers', () => {
    for (const enabledTools of [
      undefined,
      '^(list-mail-attachments|read-document)$',
      '^(list-mail-attachments)$',
      '^(get-drive-item|download-bytes)$',
    ]) {
      for (const attachmentProxy of [true, false]) {
        const { names } = registerAndCollect({
          discovery: false,
          attachmentProxy,
          httpMode: true,
          enabledTools,
        });
        const resolved = resolveRegisteredToolNames({
          readOnly: false,
          orgMode: true,
          enabledTools,
          httpMode: true,
          attachmentProxy,
        });
        const registered = [...names].filter((n) => !AUTH_TOOL_NAMES.includes(n)).sort();
        expect([...resolved].sort()).toEqual(registered);
      }
    }
  });
});

/**
 * The sentence splitter is the one fragile part: cutting at the wrong period
 * either loses true guidance or leaves half a stale sentence behind. Every
 * period shape that appears in the tips it runs over is pinned here.
 */
describe('stripStaleToolGuidance', () => {
  const suppressed = ['download-bytes', 'get-download-url'];

  it('drops only the sentences naming a suppressed tool', () => {
    const text =
      'Lists attachments. Call download-bytes with target=/x/$value. IDs come from the body.';
    expect(stripStaleToolGuidance(text, suppressed)).toBe(
      'Lists attachments. IDs come from the body.'
    );
  });

  it('leaves text untouched when nothing is suppressed or nothing matches', () => {
    const text = 'Call download-bytes with target=/x/$value.';
    expect(stripStaleToolGuidance(text, [])).toBe(text);
    expect(stripStaleToolGuidance('Lists attachments.', suppressed)).toBe('Lists attachments.');
  });

  it('does not split on an unspaced or lowercase-following period', () => {
    const text =
      'Returns @microsoft.graph.downloadUrl and ProfilePhoto.ReadWrite.All applies. ' +
      'Use get-download-url for the bytes.';
    expect(stripStaleToolGuidance(text, suppressed)).toBe(
      'Returns @microsoft.graph.downloadUrl and ProfilePhoto.ReadWrite.All applies.'
    );
  });

  it('appends the replacement once however many sentences were dropped', () => {
    const text =
      'Metadata only. Call get-download-url for large files. Call download-bytes for small ones.';
    const replacement = { text: 'Use read-document.', whenDropped: suppressed };
    expect(stripStaleToolGuidance(text, suppressed, replacement)).toBe(
      'Metadata only. Use read-document.'
    );
  });

  // A sentence dropped for an unrelated reason gets no substitute: the
  // read-document offer answers "the byte tools are gone", not "a preset
  // excluded a tool this tip cross-referenced".
  it('withholds the replacement when the dropped sentence is unrelated to it', () => {
    const text = 'Metadata only. Call list-mail-folders for the folder ids.';
    const replacement = { text: 'Use read-document.', whenDropped: ['download-bytes'] };
    expect(stripStaleToolGuidance(text, ['list-mail-folders'], replacement)).toBe('Metadata only.');
  });

  it('matches tool names on word boundaries and either separator', () => {
    // The live server's clients render these names with underscores.
    expect(mentionsToolName('call download_bytes now', ['download-bytes'])).toBe(true);
    // A longer name is not a mention of its own prefix.
    expect(mentionsToolName('use download-bytes-to-file', ['download-bytes'])).toBe(false);
    expect(mentionsToolName('use download-bytes-to-file', ['download-bytes-to-file'])).toBe(true);
  });
});
