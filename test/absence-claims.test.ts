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
 */
import { describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import type { z } from 'zod';
import {
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
const CLASSIFIED: Array<{ fragment: string; kind: 'MECHANISM' | 'GATE'; why: string }> = [
  {
    fragment: 'so no tool returns raw bytes',
    kind: 'MECHANISM',
    why: 'installResponseScrubbing strips contentBytes and base64 over 4 KB from every result, and is installed on the same condition that registers read-document.',
  },
  {
    fragment: 'nothing base64 ever enters this conversation',
    kind: 'MECHANISM',
    why: 'Same scrubber. read-document itself returns converted text, and the scrubber covers anything else that would carry document bytes.',
  },
  {
    fragment: 'The only out-of-band way to save mail attachments',
    kind: 'GATE',
    why: 'download-bytes-to-file is stdioOnly, and minting a URL requires HTTP, so wherever this string is registered no out-of-band alternative exists.',
  },
];

type Surface = { where: string; text: string };

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
  it('makes no absence claim that is not mechanism-enforced or gate-derived', () => {
    const unclassified: string[] = [];
    for (const surface of modelFacingStrings()) {
      for (const sentence of surface.text.split(/(?<=\.)\s+/)) {
        if (!ABSENCE_PHRASING.test(sentence)) continue;
        if (CLASSIFIED.some((entry) => sentence.includes(entry.fragment))) continue;
        unclassified.push(`${surface.where}: ${sentence.trim()}`);
      }
    }
    expect(unclassified).toEqual([]);
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

  it('keeps every classified entry present and justified', () => {
    // A stale allowlist is its own failure mode: an entry whose sentence no
    // longer exists hides the fact that nothing is being checked.
    const all = modelFacingStrings()
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
