/**
 * Guidance may state what is registered. It may not state what cannot be done.
 *
 * Three separate absence claims shipped and were caught one per review round --
 * "no content can be read", "no tool returns binary bytes", "no byte-level
 * fallback", plus read-document's "the ONLY way to read the content of
 * anything" -- each false as soon as a registration gate differed from the one
 * whoever wrote the sentence had in mind. Reader judgement found one of them per
 * pass, which is why this is mechanical instead: every model-facing string is
 * swept for absence phrasing, and every hit must be classified here before it
 * can ship.
 *
 * A claim of absence is legitimate in exactly two cases:
 *
 *   - MECHANISM: something enforces it. `installResponseScrubbing` strips
 *     `contentBytes` and any base64 over 4 KB from every tool result
 *     (lib/response-scrubber.ts), so "no tool returns raw bytes" is true by
 *     construction wherever that is installed.
 *   - GATE: it restates a registration gate this code is handed -- `stdioOnly`,
 *     `readOnly`, `--org-mode`.
 *
 * Everything else is TOOL-LIST-IMPLIED: it asserts something about tools the
 * string cannot see, and goes false the moment a filter, preset or flag differs.
 * Those get fixed, not allowlisted.
 *
 * Statements about what MICROSOFT GRAPH cannot do ("the default calendar cannot
 * be deleted", "cannot combine $filter with $search") are not in scope: they are
 * facts about the API, not about which tools this server registered, and no gate
 * can make them false. The phrase list below is deliberately narrow enough to
 * exclude them -- it targets superlatives and global claims about *this server*,
 * which is the shape that actually failed.
 *
 * ## What this sweep covers, and what it cannot
 *
 * Two passes, because neither alone is honest:
 *
 *   1. EMITTED. Descriptions, parameter descriptions and instructions, built by
 *      the real code across several configurations. Proves what the composed
 *      output actually says, interpolation included.
 *   2. LITERAL. Every string literal in `src/**` (minus `generated/` and
 *      `__tests__/`), read with the TypeScript scanner so comments are excluded
 *      by construction rather than by regex. This is what catches text on paths
 *      no test reaches -- the `expand_not_allowed` handler response was invisible
 *      to pass 1 and is exactly the blind spot that made a green guard read as
 *      coverage it did not have.
 *
 * Pass 2 was chosen over exercising "representative" handlers because the
 * property is *no false absence claim exists in model-facing text*, not *none is
 * emitted on the paths someone remembered to exercise* -- and a representative
 * list rots as handlers are added.
 *
 * It cannot see:
 *   - a phrase split across a template interpolation (`no ${kind} tool`): the
 *     literal spans are joined with a space, so a phrase straddling `${...}`
 *     is missed;
 *   - text assembled at runtime from variables or constants defined elsewhere;
 *   - `src/generated/client.ts`, which is Microsoft's own generated endpoint
 *     descriptions -- they reach the model but are not ours to edit, and they
 *     describe Graph rather than this server's tool set;
 *   - text originating outside this repo (the document proxy's own error
 *     messages, Graph error bodies) that a tool result may pass through.
 *
 * Logger arguments are skipped: log lines go to the operator, never to the
 * model. Startup `throw new Error` text is classified OPERATOR below for the
 * same reason, individually rather than by rule, because an Error thrown inside
 * a handler CAN reach a model and must not be waved through.
 */
import { describe, expect, it, vi } from 'vitest';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';
import ts from 'typescript';
import type { z } from 'zod';
import {
  isProxySuppressedGraphTool,
  PROXY_DOCUMENT_READ_GUIDANCE,
  resolveRegisteredToolNames,
  UTILITY_TOOLS,
} from '../src/graph-tools.js';
import { buildMcpServerInstructions } from '../src/mcp-instructions.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn(), verbose: vi.fn() },
  enableConsoleLogging: vi.fn(),
}));

/**
 * Superlatives and global absence claims about this server. Not "cannot" or
 * "unavailable": in this corpus those are overwhelmingly Graph API facts, and
 * banning them would bury the signal under 20 entries that can never go false.
 */
const ABSENCE_PHRASING =
  /only way|the only|no tool|no other|nothing|there is no|never return|no way/i;

/**
 * Each entry: a stable fragment of the sentence, why it is allowed to assert an
 * absence, and the mechanism or gate that makes it true. Adding an entry is a
 * claim that one of those two justifications applies -- if neither does, fix the
 * text instead.
 */
/*
 * MECHANISM was withdrawn from the scrubber sentences it used to bless.
 *
 * "No tool returns raw bytes" was classified MECHANISM because
 * installResponseScrubbing enforces it. It does not: proxy suppression matches
 * only GET paths ending /$value, so graph-batch (POST /$batch, no presets)
 * stays registered and can batch a GET of /me/messages/{id}/$value, and that
 * RFC 5322 body is text/plain rather than wholly valid base64, so neither
 * scrubber rule matches while its attachments ride inline. Both
 * isProxySuppressedGraphTool's docstring and response-scrubbing.ts disclose
 * this, which is what makes the classification -- not the disclosure -- the
 * defect. A written justification for a false claim is worse than an
 * unjustified one.
 *
 * The sentences now describe what the scrubber strips instead of promising what
 * cannot happen, so they no longer match ABSENCE_PHRASING and need no entry
 * here. A claim that holds only while some tool is unregistered is conditional,
 * not mechanical, and does not belong in this list.
 */
const CLASSIFIED: Array<{
  fragment: string;
  kind: 'MECHANISM' | 'GATE' | 'OPERATOR';
  why: string;
}> = [
  // Classified per sentence, not per message: each claim earns its own
  // justification, so a second sentence smuggled into an already-allowed
  // message still has to be looked at.
  {
    fragment: 'Cannot install the attachment-proxy response scrubber',
    kind: 'OPERATOR',
    why: 'Same install-time throw in response-scrubbing.ts; this is its first sentence, about the SDK handler being absent rather than about any Graph tool.',
  },
  {
    fragment: '--attachment-port requires --enable-attachment-urls',
    kind: 'OPERATOR',
    why: 'server.ts throws this during startup validation, before the MCP server exists, so no model can see it.',
  },
  {
    fragment: '--attachment-host requires --attachment-port',
    kind: 'OPERATOR',
    why: 'The neighbouring startup throw in server.ts, refused for the same reason and equally unreachable from a tool call.',
  },
  {
    fragment: 'The only out-of-band way to save mail attachments',
    kind: 'GATE',
    why: 'download-bytes-to-file is stdioOnly, and minting a URL requires HTTP, so wherever this string is registered no out-of-band alternative exists.',
  },
];

type Surface = { where: string; text: string };

/** Every .ts file under src/, minus generated code and unit tests. */
function sourceFiles(dir: string, out: string[] = []): string[] {
  for (const entry of readdirSync(dir)) {
    const full = path.join(dir, entry);
    if (statSync(full).isDirectory()) {
      if (entry === 'generated' || entry === '__tests__') continue;
      sourceFiles(full, out);
    } else if (entry.endsWith('.ts')) {
      out.push(full);
    }
  }
  return out;
}

/** True when this literal is an argument to a logger call, which the model never sees. */
function insideLoggerCall(node: ts.Node): boolean {
  for (let current = node.parent; current; current = current.parent) {
    if (ts.isCallExpression(current)) {
      const target = current.expression.getText();
      if (/^logger\./.test(target) || /^console\./.test(target)) return true;
    }
  }
  return false;
}

/**
 * Pass 2: string literals as written, via the TypeScript scanner. Sees handler
 * responses and any other branch no test reaches; see the file docstring for
 * what it cannot see.
 */
function literalStrings(): Surface[] {
  const surfaces: Surface[] = [];
  const root = path.join(import.meta.dirname, '..', 'src');

  const isConcat = (node: ts.Node): node is ts.BinaryExpression =>
    ts.isBinaryExpression(node) && node.operatorToken.kind === ts.SyntaxKind.PlusToken;

  /**
   * A whole string expression as one text, so a sentence is judged whole.
   *
   * These strings are written as `'...' + '...' + '...'` across many lines, and
   * judging each fragment on its own split sentences down the middle: the
   * classified "so no tool returns raw bytes" landed in one literal and the rest
   * of its sentence in the next, so a legitimate claim read as an unclassified
   * one. Non-literal operands and `${...}` gaps become a single space, which is
   * why a phrase straddling an interpolation is still invisible.
   */
  const flatten = (node: ts.Node): string => {
    if (ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node)) return node.text;
    if (ts.isTemplateExpression(node)) {
      return [node.head.text, ...node.templateSpans.map((span) => span.literal.text)].join(' ');
    }
    if (isConcat(node)) return flatten(node.left) + flatten(node.right);
    if (ts.isParenthesizedExpression(node)) return flatten(node.expression);
    return ' ';
  };

  const holdsLiteral = (node: ts.Node): boolean =>
    ts.isStringLiteral(node) ||
    ts.isNoSubstitutionTemplateLiteral(node) ||
    ts.isTemplateExpression(node) ||
    (isConcat(node) && (holdsLiteral(node.left) || holdsLiteral(node.right))) ||
    (ts.isParenthesizedExpression(node) && holdsLiteral(node.expression));

  for (const file of sourceFiles(root)) {
    const source = ts.createSourceFile(
      file,
      readFileSync(file, 'utf8'),
      ts.ScriptTarget.Latest,
      true
    );
    const visit = (node: ts.Node) => {
      // Only the outermost node of a string expression, so concatenated
      // fragments are never judged on their own.
      const parent = node.parent;
      const isInner =
        (parent && isConcat(parent)) || (parent && ts.isParenthesizedExpression(parent));
      if (!isInner && holdsLiteral(node)) {
        if (!insideLoggerCall(node)) {
          const line = source.getLineAndCharacterOfPosition(node.getStart()).line + 1;
          surfaces.push({ where: `${path.relative(root, file)}:${line}`, text: flatten(node) });
        }
        return;
      }
      ts.forEachChild(node, visit);
    };
    visit(source);
  }
  return surfaces;
}

/** Every string this server hands the model, across the configurations that shape them. */
function modelFacingStrings(): Surface[] {
  const surfaces: Surface[] = [];
  const root = path.join(import.meta.dirname, '..');

  const endpoints = JSON.parse(
    readFileSync(path.join(root, 'src', 'endpoints.json'), 'utf8')
  ) as Array<Record<string, string>>;
  for (const endpoint of endpoints) {
    for (const field of ['llmTip', 'descriptionOverride'] as const) {
      if (endpoint[field]) {
        surfaces.push({
          where: `endpoints.json:${endpoint.toolName}:${field}`,
          text: endpoint[field],
        });
      }
    }
  }

  for (const utility of UTILITY_TOOLS) {
    surfaces.push({ where: `utility:${utility.name}:description`, text: utility.description });
    const shape = utility.buildSchema({
      multiAccount: true,
      accountNames: ['user@example.com'],
    } as never);
    for (const [param, zodType] of Object.entries(shape)) {
      const description = (zodType as z.ZodTypeAny).description;
      if (description) {
        surfaces.push({ where: `utility:${utility.name}:param:${param}`, text: description });
      }
    }
  }

  surfaces.push({ where: 'PROXY_DOCUMENT_READ_GUIDANCE', text: PROXY_DOCUMENT_READ_GUIDANCE });

  const configurations: Array<[string, Parameters<typeof resolveRegisteredToolNames>[0]]> = [
    ['proxy', { orgMode: true, httpMode: true, attachmentProxy: true }],
    ['stdio', { orgMode: true }],
    ['http', { orgMode: true, httpMode: true }],
    ['readonly-proxy', { orgMode: true, httpMode: true, attachmentProxy: true, readOnly: true }],
    ['narrow-graph-only', { orgMode: true, enabledTools: '^get-mail-message$' }],
    [
      'proxy-no-read-document',
      {
        orgMode: true,
        httpMode: true,
        attachmentProxy: true,
        enabledTools: '^list-mail-attachments$',
      },
    ],
  ];
  for (const [label, gates] of configurations) {
    for (const discovery of [false, true]) {
      surfaces.push({
        where: `instructions:${label}:discovery=${discovery}`,
        text: buildMcpServerInstructions({
          orgMode: true,
          readOnly: Boolean(gates.readOnly),
          multiAccount: true,
          discovery,
          registeredTools: resolveRegisteredToolNames(gates),
        }),
      });
    }
  }
  return surfaces;
}

describe('absence claims in model-facing guidance', () => {
  const unclassifiedIn = (surfaces: Surface[]): string[] => {
    const unclassified: string[] = [];
    for (const surface of surfaces) {
      for (const sentence of surface.text.split(/(?<=\.)\s+/)) {
        if (!ABSENCE_PHRASING.test(sentence)) continue;
        if (CLASSIFIED.some((entry) => sentence.includes(entry.fragment))) continue;
        unclassified.push(`${surface.where}: ${sentence.trim()}`);
      }
    }
    return unclassified;
  };

  it('makes no absence claim that is not mechanism-enforced or gate-derived', () => {
    expect(unclassifiedIn(modelFacingStrings())).toEqual([]);
  });

  // Pass 2. Without this the guard was green on text it had never looked at.
  it('makes no unclassified absence claim in any source string literal', () => {
    expect(unclassifiedIn(literalStrings())).toEqual([]);
  });

  /*
   * Proof of coverage, not just of greenness. A sweep that silently stopped
   * collecting -- a moved file, a changed directory layout, a scanner that
   * returns nothing -- would pass every assertion above by finding nothing to
   * check. These pin the two things the literal pass exists to reach.
   */
  it('actually reaches handler response text the emitted pass cannot', () => {
    const literals = literalStrings();
    const emitted = modelFacingStrings()
      .map((surface) => surface.text)
      .join('\n');

    // The expand_not_allowed refusal: a tool-result string on a branch no
    // description or instruction pass can produce.
    const expandGuard = literals.filter((surface) =>
      surface.text.includes(
        'strips contentBytes fields and large base64 values out of tool results'
      )
    );
    expect(expandGuard.length).toBeGreaterThan(0);
    expect(expandGuard[0].where).toMatch(/^graph-tools\.ts:\d+$/);
    expect(emitted).not.toContain(
      'strips contentBytes fields and large base64 values out of tool results'
    );

    // And the sweep is looking at a real corpus, not an empty one.
    expect(literals.length).toBeGreaterThan(500);
    expect(new Set(literals.map((s) => s.where.split(':')[0])).size).toBeGreaterThan(10);
  });

  it('excludes logger arguments, which reach the operator and not the model', () => {
    const literals = literalStrings();
    // server.ts warns about a filtered-out read-document; that text says
    // "nothing replaced them" and is a log line, so it must not be swept.
    expect(literals.some((surface) => surface.text.includes('nothing replaced them'))).toBe(false);
  });

  // The two the reviews caught, pinned by name so they cannot come back under a
  // reworded sentence that the phrase sweep happens not to match.
  it('does not claim read-document is the only way to read content', () => {
    const readDocument = UTILITY_TOOLS.find((utility) => utility.name === 'read-document')!;
    expect(readDocument.description).not.toMatch(/only way/i);
    expect(readDocument.description).not.toMatch(/ONLY/);
    // The positive guidance survives.
    expect(readDocument.description).toContain('Read any Microsoft 365 document as markdown');
  });

  it('does not claim there is no fallback when read-document refuses', () => {
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
    expect(instructions).not.toMatch(/no byte-level fallback/i);
    expect(instructions).not.toMatch(/there is no/i);
    // The refusal itself is still described -- that is a guard, not an absence.
    expect(instructions).toContain('identity_not_supported');
    expect(instructions).toMatch(/OAuth, OBO, or bearer mode/);
  });

  it('qualifies every direction to read-document with the identity limit', () => {
    // read-document is registered and visible in tools/list in OAuth, OBO and
    // bearer deployments, and refuses every call there
    // (graph-tools.ts: isOAuthModeEnabled() || getRequestTokens()). Anything
    // that sends the model to it has to name that condition, not just the
    // instructions -- the tool description and the guidance appended to Graph
    // tool descriptions are separate surfaces and were missed once each.
    const readDocument = UTILITY_TOOLS.find((utility) => utility.name === 'read-document')!;
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

    for (const [where, text] of [
      ['read-document description', readDocument.description],
      ['PROXY_DOCUMENT_READ_GUIDANCE', PROXY_DOCUMENT_READ_GUIDANCE],
      ['instructions', instructions],
    ] as Array<[string, string]>) {
      expect(text, `${where} must name the identity limit`).toContain('identity_not_supported');
      expect(text, `${where} must name the modes`).toMatch(/OAuth, OBO, or bearer mode/);
    }
  });

  /*
   * The scrubber's guarantee is narrower than "no tool returns raw bytes", and
   * that sentence was classified MECHANISM on the strength of it -- a written
   * justification for a false claim, which is worse than the unclassified claim
   * it replaced.
   *
   * The premise, from this repo rather than from our deployment: graph-batch is
   * a real tool (endpoints.json, POST /$batch, `scopes: []`, no `presets`), and
   * proxy suppression only matches GET paths ending /$value
   * (isProxySuppressedGraphTool), so nothing keeps it out of a proxy server that
   * runs without a preset or filter. A batched GET of /me/messages/{id}/$value
   * returns RFC 5322 source, which is text/plain and not wholly valid base64, so
   * neither scrubber rule matches it while its attachments ride inline as base64
   * -- documented at graph-tools.ts (isProxySuppressedGraphTool's docstring) and
   * again at response-scrubbing.ts, which says the wrapper "cannot see into" it.
   *
   * So the text describes what the mechanism does instead of promising what
   * cannot happen. Describing a mechanism is verifiable; promising an absence is
   * the thing that has now failed repeatedly.
   */
  it('registers graph-batch under proxy mode with no filter, so the premise holds', () => {
    const registered = resolveRegisteredToolNames({
      orgMode: true,
      httpMode: true,
      attachmentProxy: true,
    });
    expect(registered.has('graph-batch')).toBe(true);
    // And the suppression rule genuinely cannot reach it.
    expect(isProxySuppressedGraphTool('post', '/$batch')).toBe(false);
    expect(isProxySuppressedGraphTool('get', '/me/messages/{message-id}/$value')).toBe(true);
  });

  it('promises no byte-free tool result, and describes the scrubber instead', () => {
    const readDocument = UTILITY_TOOLS.find((utility) => utility.name === 'read-document')!;
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

    // Every copy of the over-promise, including the handler response.
    const corpus = [
      ['read-document description', readDocument.description],
      ['PROXY_DOCUMENT_READ_GUIDANCE', PROXY_DOCUMENT_READ_GUIDANCE],
      ['instructions', instructions],
      ...literalStrings().map((s) => [s.where, s.text] as [string, string]),
    ] as Array<[string, string]>;
    for (const [where, text] of corpus) {
      expect(text, `${where} promises byte-free results`).not.toMatch(/no tool returns raw bytes/i);
      expect(text, `${where} promises no base64 anywhere`).not.toMatch(/nothing base64/i);
    }

    // Replaced by a description of what is actually stripped.
    for (const [where, text] of [
      ['PROXY_DOCUMENT_READ_GUIDANCE', PROXY_DOCUMENT_READ_GUIDANCE],
      ['instructions', instructions],
    ] as Array<[string, string]>) {
      expect(text, `${where} must name the field`).toContain('contentBytes');
      expect(text, `${where} must name the shape rule`).toMatch(/base64/);
    }
  });

  it('keeps every classified entry present and justified', () => {
    // A stale allowlist is its own failure mode: an entry whose sentence no
    // longer exists hides the fact that nothing is being checked.
    const all = [...modelFacingStrings(), ...literalStrings()]
      .map((surface) => surface.text)
      .join('\n');
    for (const entry of CLASSIFIED) {
      expect(all, `classified entry no longer present: ${entry.fragment}`).toContain(
        entry.fragment
      );
      expect(entry.why.length).toBeGreaterThan(40);
    }
  });
});
