import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { randomUUID } from 'crypto';
import logger from './logger.js';
import { auditLog, getUserIdentityForAudit } from './audit-log.js';
import GraphClient from './graph-client.js';
import { isDestructiveOperation } from './lib/destructive-ops.js';
import { describePathParam } from './lib/path-params.js';
import { getAttachmentMinting } from './lib/attachment-minting.js';
import { getAttachmentProxy } from './lib/attachment-proxy-runtime.js';
import {
  buildAttachmentUrl,
  MAX_REDEMPTIONS,
  TicketStoreFullError,
} from './lib/attachment-tickets.js';
import AuthManager, {
  getEndpointScopeGroups,
  getMissingAllowedScopesForGroups,
  parseAllowedScopes,
} from './auth.js';
import { api } from './generated/client.js';
import { api as betaApi } from './generated/client-beta.js';

// Tools from every Graph API version share one registry. Each tool's version is carried
// by its endpoints.json config (apiVersion), so the generated clients stay version-agnostic
// and the runtime picks the URL prefix per request. v1.0 endpoints are unchanged.
const allEndpoints = [...api.endpoints, ...betaApi.endpoints];
import { z } from 'zod';
import { readFileSync } from 'fs';
import { access } from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import { getCategoryPattern, TOOL_CATEGORIES } from './tool-categories.js';
import { getRequestTokens } from './request-context.js';
import { parseTeamsUrl } from './lib/teams-url-parser.js';
import { buildBM25Index, scoreQuery, tokenize, type BM25Index } from './lib/bm25.js';
import { deriveTargetResource, type AuditTargetResource } from './audit-target-resource.js';
export interface DiscoverySearchIndex {
  bm25: BM25Index;
  nameTokens: Map<string, Set<string>>;
}
import { describeToolSchema, describeUtilityToolSchema } from './lib/tool-schema.js';
import {
  TOP_UNSUPPORTED_DELTA_TOOLS,
  shouldOmitTopParam,
  paginationAllowed,
  positiveIntFromEnv,
  DEFAULT_MAX_PAGES,
  getMaxPages,
  isFetchAllPagesApplicable,
  FILTER_PARAM_DESCRIPTION,
  SEARCH_PARAM_DESCRIPTION,
  SELECT_PARAM_DESCRIPTION,
  EXPAND_PARAM_DESCRIPTION,
  ORDERBY_PARAM_DESCRIPTION,
  TOP_PARAM_DESCRIPTION,
  SKIP_PARAM_DESCRIPTION,
  COUNT_PARAM_DESCRIPTION,
  CONFIRM_PARAM_DESCRIPTION,
  TIMEZONE_PARAM_DESCRIPTION,
  EXPAND_EXTENDED_PROPERTIES_PARAM_DESCRIPTION,
  getAccountParamDescription,
  getFetchAllPagesParamDescription,
} from './lib/param-descriptions.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

interface EndpointConfig {
  pathPattern: string;
  method: string;
  toolName: string;
  scopes?: string[] | string[][];
  workScopes?: string[] | string[][];
  apiVersion?: string; // Graph API version ('v1.0' default, or 'beta'). Selects spec + URL prefix.
  returnDownloadUrl?: boolean;
  supportsTimezone?: boolean;
  supportsExpandExtendedProperties?: boolean;
  llmTip?: string;
  // Replaces the Microsoft-supplied base description everywhere it is surfaced (tool
  // registration, BM25 discovery index, search-tools, get-tool-schema). Use when the
  // generated description leads with the wrong Graph operation. llmTip is still appended after.
  descriptionOverride?: string;
  skipEncoding?: string[]; // Parameter names that should NOT be URL-encoded (for function-style API calls)
  contentType?: string;
  acceptType?: string; // Custom Accept header for endpoints returning non-JSON content (e.g., text/vtt)
  readOnly?: boolean; // When true, allow this endpoint in read-only mode even if method is not GET
  presets?: string[]; // Presets this endpoint belongs to (mail, outlook, personal, ...)
  // JSON Schema for the request body of an endpoint that Microsoft has NOT published
  // in its OpenAPI metadata. Consumed at generate time by bin/modules/simplified-openapi.mjs
  // to synthesize a typed requestBody (instead of a generic object), so the generated client
  // exposes a validated `body` param. Ignored for endpoints already present in the spec.
  requestBodySchema?: Record<string, unknown>;
}

const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, 'endpoints.json'), 'utf8')
) as EndpointConfig[];

/**
 * Prefix beta-version tools with a [beta] marker so the instability is visible in the
 * tool description itself, regardless of what (if anything) the llmTip says. Tools on
 * v1.0 (the default) are returned unchanged.
 */
function withApiVersionPrefix(description: string, config?: EndpointConfig): string {
  return config?.apiVersion === 'beta' ? `[beta] ${description}` : description;
}

/** When set to a positive integer, caps Graph `$top` on list requests (see README). */
function maxTopFromEnv(): number | undefined {
  const raw = process.env.MS365_MCP_MAX_TOP;
  if (raw === undefined || raw === '') return undefined;
  const n = Number.parseInt(raw, 10);
  if (!Number.isFinite(n) || n < 1) {
    logger.warn(
      `Ignoring invalid MS365_MCP_MAX_TOP=${JSON.stringify(raw)} (use a positive integer)`
    );
    return undefined;
  }
  return n;
}

function clampTopQueryParam(queryParams: Record<string, string>): void {
  const cap = maxTopFromEnv();
  if (cap === undefined || queryParams['$top'] === undefined) return;
  const requested = Number.parseInt(queryParams['$top'], 10);
  if (!Number.isFinite(requested) || requested <= cap) return;
  logger.info(`Clamping $top from ${requested} to ${cap} (MS365_MCP_MAX_TOP)`);
  queryParams['$top'] = String(cap);
}

// Outlook mail lives under a mailbox owner. /chats, /teams and /planner also have
// /messages collections, and directory search has neither prefix — none of them share
// mail's quoting convention, so all are left alone.
const MAILBOX_OWNER_PATH = /^\/(?:me|users\/[^/]+)\//i;
const MAIL_COLLECTION_PATH = /\/(?:messages|mailFolders)(?:\/|$)/i;

function isOutlookMailPath(path: string): boolean {
  return MAILBOX_OWNER_PATH.test(path) && MAIL_COLLECTION_PATH.test(path);
}

/** A quoted run starting at `start` (an opening quote), with escapes preserved. */
function readQuotedSegment(
  expr: string,
  start: number
): { segment: string; end: number } | undefined {
  let j = start + 1;
  let segment = '';
  while (j < expr.length) {
    if (expr[j] === '\\' && expr[j + 1] === '"') {
      segment += '\\"';
      j += 2;
      continue;
    }
    if (expr[j] === '"') return { segment, end: j };
    segment += expr[j];
    j += 1;
  }
  return undefined;
}

// The properties KQL recognises on a message. Shape alone is not enough to tell a clause
// from a phrase: "RE: quarterly report" and "Q3: plan.pdf" both look like property:value.
const MAIL_SEARCH_PROPERTIES = new Set([
  'attachment',
  'bcc',
  'body',
  'category',
  'cc',
  'from',
  'hasattachment',
  'hasattachments',
  'importance',
  'isread',
  'kind',
  'participants',
  'received',
  'recipients',
  'sent',
  'size',
  'subject',
  'to',
]);

/** `property:` or a comparison — `received>=2024-01-01`, `size>1000`. */
const CLAUSE_HEAD = /^([A-Za-z]+)\s*(?::|<=|>=|<>|=|<|>)/;

/**
 * True only for a quoted run that really is a clause. Anything else — including a phrase
 * that merely opens with a word and a colon — keeps its grouping quotes. Erring this way
 * leaves an unrecognised property unrepaired rather than silently changing what a valid
 * phrase search means.
 */
function isPropertyClause(segment: string): boolean {
  const head = CLAUSE_HEAD.exec(segment);
  return head ? MAIL_SEARCH_PROPERTIES.has(head[1].toLowerCase()) : false;
}

/**
 * Rewrite the interior of a mail KQL expression so it can be wrapped in one pair of
 * double quotes.
 *
 * A quoted run is either a phrase, whose quotes group the words and must survive
 * (escaped as \"), or a whole clause the caller quoted by mistake, whose quotes must go.
 * Two signals separate them: a run introduced by `property:` is always a phrase, even
 * when its own text contains a colon (subject:"RE: quarterly report"); otherwise a run
 * that itself starts with `property:` is the mistake ("from:john" AND subject:meeting).
 */
function rewriteMailSearchQuotes(expr: string): string {
  let out = '';
  let i = 0;
  while (i < expr.length) {
    if (expr[i] === '\\' && expr[i + 1] === '"') {
      out += '\\"';
      i += 2;
      continue;
    }
    if (expr[i] !== '"') {
      out += expr[i];
      i += 1;
      continue;
    }
    const run = readQuotedSegment(expr, i);
    if (!run) {
      // Unbalanced quote: keep the text, drop the stray delimiter.
      out += expr.slice(i + 1);
      break;
    }
    const introducedByProperty = i > 0 && expr[i - 1] === ':';
    const isPhrase = introducedByProperty || !isPropertyClause(run.segment);
    out += isPhrase ? `\\"${run.segment}\\"` : run.segment;
    i = run.end + 1;
  }
  return out.trim();
}

/**
 * Outlook mail wants the whole KQL expression inside one pair of double quotes
 * ($search="from:x AND subject:y"). Models quote each clause instead
 * ($search='"from:x" AND subject:y'), or send a phrase with no enclosing pair
 * ($search='subject:"quarterly report"'). Both are 400s. Normalize to one enclosing
 * pair, mirroring the Body auto-wrap already done in executeGraphTool.
 *
 * Verified against Graph: a bare single term and a correctly wrapped expression both
 * succeed; 'subject:"quarterly report"' and '"quarterly report" AND from:x' are both
 * rejected until the enclosing pair is added.
 */
function normalizeSearchQueryParam(queryParams: Record<string, string>, path: string): void {
  if (!isOutlookMailPath(path)) return;

  const raw = queryParams['$search'];
  if (raw === undefined) return;
  const trimmed = raw.trim();

  // Nothing searchable — Graph rejects it, and sending it cannot be what was meant.
  if (trimmed === '' || /^["'\s]+$/.test(trimmed)) {
    delete queryParams['$search'];
    logger.warn("Dropping empty '$search' parameter");
    return;
  }

  // An expression already inside one enclosing pair is unwrapped first, so its interior
  // is judged on its own terms and re-wrapped unchanged. Without this, a correctly
  // wrapped free-text search ("quarterly report") would be read as a phrase and become
  // a phrase search ("\"quarterly report\"").
  let expr = trimmed;
  if (expr.startsWith('"')) {
    const whole = readQuotedSegment(expr, 0);
    if (whole && whole.end === expr.length - 1) expr = whole.segment;
  }

  const inner = rewriteMailSearchQuotes(expr);
  if (inner === '') return;
  const normalized = `"${inner}"`;
  if (normalized !== raw) {
    logger.info(`Auto-corrected parameter '$search': normalized KQL quoting to ${normalized}`);
    queryParams['$search'] = normalized;
  }
}

const DEFAULT_MAX_ITEMS = 10_000;

// Canonical definitions of TOP_UNSUPPORTED_DELTA_TOOLS, paginationAllowed, and
// positiveIntFromEnv live in lib/param-descriptions.ts so tool-schema.ts can use
// them without circling back through graph-tools.ts, and so the description text
// they parameterize can't drift between the two registration paths (see that
// file's header comment).

// Canonical definition lives in lib/destructive-ops.ts so tool-schema.ts can
// use it without circling back through graph-tools.ts; re-exported here for
// external callers (tests, etc.) that imported it from this module.
export { isDestructiveOperation };

/**
 * Defense-in-depth: destructive tools require an explicit `confirm: true` from
 * the caller before they reach Microsoft Graph. Mitigates accidental
 * sendMail / deleteEvent / etc. when an LLM misroutes a request or follows an
 * injected instruction. Opt in per-deployment via MS365_MCP_REQUIRE_CONFIRM=true
 * (default off, so the gate is a non-breaking, additive opt-in that can coexist
 * with client-side elicitation prompts).
 */
function isConfirmGateEnabled(): boolean {
  return process.env.MS365_MCP_REQUIRE_CONFIRM === 'true';
}

type TextContent = {
  type: 'text';
  text: string;
  [key: string]: unknown;
};

type ImageContent = {
  type: 'image';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type AudioContent = {
  type: 'audio';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type ResourceTextContent = {
  type: 'resource';
  resource: {
    text: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceBlobContent = {
  type: 'resource';
  resource: {
    blob: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceContent = ResourceTextContent | ResourceBlobContent;

type ContentItem = TextContent | ImageContent | AudioContent | ResourceContent;

interface CallToolResult {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;

  [key: string]: unknown;
}

interface UtilityToolContext {
  graphClient: GraphClient;
  authManager?: AuthManager;
  multiAccount: boolean;
  accountNames: string[];
}

interface UtilityTool {
  name: string;
  // Synthetic for display in search-tools / get-tool-schema. The `tool:` prefix
  // marks these as non-Graph so an LLM doesn't try to construct a Graph URL from them.
  method: string;
  path: string;
  description: string;
  searchKeywords?: string;
  buildSchema: (ctx: UtilityToolContext) => Record<string, z.ZodTypeAny>;
  execute: (params: Record<string, unknown>, ctx: UtilityToolContext) => Promise<CallToolResult>;
  readOnlyHint?: boolean;
  openWorldHint?: boolean;
  // When true, this tool writes to the server's local filesystem and is only
  // registered in stdio mode — never in HTTP/OAuth mode, where a remote client
  // must not be able to write arbitrary files onto the host.
  stdioOnly?: boolean;
  // Registered ONLY under --attachment-proxy. Without a configured proxy the
  // tool has nothing to call, so registering it would advertise a capability
  // that answers an error to every invocation.
  proxyOnly?: boolean;
  // Puts raw resource bytes into the model's context, so NOT registered under
  // --attachment-proxy. download-bytes-to-file is deliberately unmarked: its
  // bytes go to a local file and never to the model, and it is stdio-only
  // anyway, which proxy mode never is.
  bytesToModel?: boolean;
}

interface DisabledToolScope {
  toolName: string;
  missingScopes: string[];
}

function formatDisabledToolsForLog(disabledTools: DisabledToolScope[]): string {
  const shown = disabledTools
    .slice(0, 20)
    .map((tool) => `${tool.toolName} (missing: ${tool.missingScopes.join(', ')})`);
  const suffix =
    disabledTools.length > shown.length ? `, ... +${disabledTools.length - shown.length} more` : '';
  return `${shown.join('; ')}${suffix}`;
}

/**
 * In OAuth/HTTP bearer mode the `account` parameter cannot switch identities —
 * every Graph call uses the connecting client's bearer token. Previously a
 * provided `account` was silently ignored and the bearer user's data returned
 * (discussion #467). Returns an error message when an `account` param is
 * provided that the bearer identity cannot honor; a param matching the bearer's
 * own identity passes through. Returns null when account routing via the MSAL
 * cache is available (stdio mode, or HTTP with --trust-proxy-auth).
 */
async function checkAccountParamInBearerMode(
  accountParam: string | undefined,
  authManager?: AuthManager
): Promise<string | null> {
  if (!accountParam || !authManager) return null;
  const contextToken = getRequestTokens()?.accessToken;
  if (!contextToken && !authManager.isOAuthModeEnabled()) return null;
  const bearerToken = contextToken ?? (await authManager.getToken().catch(() => null)) ?? undefined;
  const bearerIdentity = getUserIdentityForAudit(bearerToken);
  if (bearerIdentity && bearerIdentity.toLowerCase() === accountParam.toLowerCase()) return null;
  return (
    `The 'account' parameter is not supported in HTTP/OAuth mode: every request uses the identity ` +
    `of the connecting client's bearer token` +
    (bearerIdentity ? ` ('${bearerIdentity}')` : '') +
    `, so account switching is not possible. To act as '${accountParam}', reconnect the MCP client ` +
    `authenticated as that account, or run the server in stdio mode (or HTTP with --trust-proxy-auth) ` +
    `where cached accounts are available.`
  );
}

// The Graph byte targets that have no pre-authenticated URL of their own, and therefore the exact
// set --enable-attachment-urls exists to serve. Each is matched separately in get-download-url
// because each gets its own refusal message when the flag is off.
//
// Named and exported rather than inlined at the three call sites so that "which resources does this
// flag serve?" has one answer in the codebase. Preset membership for get-download-url is derived
// from this set (tool-categories.ts), and its tests re-derive it from these patterns rather than
// restating them -- a copy would let the flag's reach grow while the preset wiring stayed still,
// which is precisely how the tool came to be missing from the mail-shaped presets.
//
// Deliberately narrow: match only real Graph mail/calendar attachment resources so driveItem path
// addressing with folders literally named messages/events/attachments is not falsely rejected.
const MAIL_EVENT_ATTACHMENT_TARGET =
  /^(?:\/me|\/users\/[^/]+|\/groups\/[^/]+)\/(?:messages|events)\/[^/]+\/attachments\//;
const MEETING_RECORDING_TARGETS = [
  /^(?:\/me|\/users\/[^/]+)\/onlineMeetings\/[^/]+\/recordings\/[^/]+(?:\/content)?$/,
  /^\/communications\/calls\/[^/]+\/recordings\/[^/]+(?:\/content)?$/,
];
const VALUE_BYTE_TARGET = /\/\$value$/;

export const MINTABLE_TARGET_PATTERNS: readonly RegExp[] = [
  MAIL_EVENT_ATTACHMENT_TARGET,
  ...MEETING_RECORDING_TARGETS,
  VALUE_BYTE_TARGET,
];

/**
 * Mint a server-served download URL for a Graph byte resource Graph itself
 * exposes no pre-authenticated URL for, or return null if minting is off.
 *
 * **This grants no authority the calling agent did not already hold.** Every
 * target that reaches here is one `download-bytes` would fetch for the same
 * caller on the same account; the ticket only moves those bytes out of the
 * agent's context window and into a direct transfer. That is the whole
 * security argument for the feature, and it is why minting is scoped to the
 * byte endpoints below rather than to any Graph path.
 *
 * Returns null when the feature is disabled, so the caller falls through to
 * the refusal it would have given before -- the tool's behaviour is unchanged
 * for anyone not running with `--enable-attachment-urls`.
 */
async function mintDownloadUrl(
  target: string,
  accountParam: string | undefined,
  authManager: AuthManager | undefined
): Promise<CallToolResult | null> {
  const minting = getAttachmentMinting();
  if (!minting) return null;

  // Refuse whenever this request's Graph identity comes from the caller rather
  // than from this server's own token cache.
  //
  // **Both halves of this predicate are load-bearing, and checking only the
  // first is an authority escalation, not merely a broken feature.**
  // `isOAuthModeEnabled()` is true only for MS365_MCP_OAUTH_TOKEN and the
  // oauth-provider path; it is *false* in plain `--http` bearer mode and in
  // `--obo`, both of which still run the tool inside a request context holding
  // the caller's token. In those modes `download-bytes` reads as the caller
  // while a redeemed ticket reads as whatever account this server has cached --
  // so minting would let a caller ask under one identity and have the bytes
  // fetched under another. Every other token site in this file pairs these two
  // checks (see the `getRequestTokens()` guards below); this one must too.
  if (authManager?.isOAuthModeEnabled() || getRequestTokens()) {
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            error:
              'Server-minted download URLs are unavailable when Graph identity comes from the request (OAuth, OBO, or bearer mode): the URL is redeemed later without an Authorization header, so the bytes would be fetched as a different identity than the one that asked for them. Use download-bytes.',
          }),
        },
      ],
      isError: true,
    };
  }

  // Validated here rather than left to the caller further down: the three
  // early mint sites return before the tool reaches its own account check, so
  // without this an unusable `account` would be baked into a ticket and only
  // surface as a 502 at redemption, long after the agent could act on it.
  const accountModeError = await checkAccountParamInBearerMode(accountParam, authManager);
  if (accountModeError) {
    return {
      content: [{ type: 'text', text: JSON.stringify({ error: accountModeError }) }],
      isError: true,
    };
  }

  let ticket: { id: string; expiresAtMs: number };
  try {
    ticket = minting.store.mint(target, accountParam);
  } catch (error) {
    if (error instanceof TicketStoreFullError) {
      return {
        content: [{ type: 'text', text: JSON.stringify({ error: error.message }) }],
        isError: true,
      };
    }
    throw error;
  }

  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify({
          downloadUrl: buildAttachmentUrl(minting.config, ticket.id),
          expiresAt: new Date(ticket.expiresAtMs).toISOString(),
          maxFetches: MAX_REDEMPTIONS,
          // Stated as an explicit false rather than dropped. The field used to
          // say `true`, and both the agent skills written against this tool and
          // any model that learned the old shape look for it; an absent field
          // reads as "unknown, assume the old rule", while `false` contradicts
          // it outright.
          singleUse: false,
          note:
            `Served by this server, not by Microsoft Graph. This URL accepts up to ${MAX_REDEMPTIONS} ` +
            `fetches until expiresAt -- enough to probe a document and then convert it, with one ` +
            `attempt to spare. Every fetch counts, including one that fails, so reuse this same URL ` +
            `for a retry or a continuation instead of minting another.`,
          // The recovery an agent could not previously work out: a converter
          // reports its own `fetch_failed`/404 with nothing to say whether the
          // URL is retryable, dead, or was never valid.
          onFetchFailure:
            `Fetch the same downloadUrl again -- a failed fetch does not invalidate it. Only a 404 ` +
            `means it is finished (all ${MAX_REDEMPTIONS} fetches used, or expiresAt passed); call ` +
            `get-download-url again for a fresh URL in that case, and only in that case.`,
        }),
      },
    ],
  };
}

/** The three facts every read-document failure carries, known or not. */
interface AttachmentFacts {
  name: string | null;
  contentType: string | null;
  size: number | null;
}

const UNKNOWN_ATTACHMENT: AttachmentFacts = { name: null, contentType: null, size: null };

/** Fixed text standing in for a redacted ticket URL or ticket id. */
const REDACTED_ATTACHMENT_URL = '<attachment url redacted>';

/**
 * Strip a minted ticket's live credential out of proxy-supplied text before
 * any of it can reach the model.
 *
 * The proxy is handed the signed ticket URL as its `uri` argument, and a
 * generic converter's error text ordinarily echoes back the address it
 * failed to fetch ("could not reach <uri>: 404") -- an entirely unremarkable
 * failure shape, not a hostile one, and none of this server's own code
 * chooses that text. Whatever the proxy sends back in `message` (or, in
 * principle, in `markdown`) is otherwise returned to the caller verbatim, so
 * without this the ticket URL -- carrying a live, redeemable credential --
 * would land in the model's context exactly where this feature exists to
 * keep it out.
 *
 * The full URL is stripped first (so a clean echo collapses to one
 * placeholder instead of the id and the surrounding query both vanishing
 * separately), then the bare ticket id is stripped on its own, because that
 * id is the actual credential and the rest of the URL is not: the redemption
 * route (`attachment-route.ts`) authorises solely on `t`, ignoring
 * `dgk`/`dgx`/`dgs` entirely, so a URL missing every parameter except a live
 * `t` is exactly as dangerous as the whole thing. Matching the id as a bare
 * substring -- not only inside the full URL -- also catches a proxy that
 * echoes the URL truncated at a delimiter, percent-encoded, or with its query
 * reordered or mangled: percent-encoding only escapes characters outside
 * `[A-Za-z0-9_-]`, and a ticket id is entirely within that set, so it
 * survives every one of those transformations unchanged and a plain string
 * search still finds it.
 */
function redactAttachmentSecrets(text: string, ticketId: string, ticketUrl: string): string {
  return text
    .split(ticketUrl)
    .join(REDACTED_ATTACHMENT_URL)
    .split(ticketId)
    .join(REDACTED_ATTACHMENT_URL);
}

/**
 * One error shape for every read-document failure.
 *
 * `name`/`contentType`/`size` are always present, null included. An agent that
 * cannot read a document can still tell the user what it saw -- "a 195 KB PDF
 * called report.pdf that the converter refused" is an answer; "an error" is not.
 * Present-and-null rather than omitted, because an absent key reads to a model
 * as "not applicable" instead of "not known".
 */
function readDocumentError(
  code: string,
  message: string,
  attachment: AttachmentFacts,
  proxyCode?: string
): CallToolResult {
  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify({
          error: code,
          ...(proxyCode ? { proxyCode } : {}),
          message,
          name: attachment.name,
          contentType: attachment.contentType,
          size: attachment.size,
        }),
      },
    ],
    isError: true,
  };
}

/**
 * Error codes this server promises.
 *
 * Anything the proxy says that is not in here is surfaced as `proxy_error` with
 * the proxy's own code beside it, unaltered. That is the honest consequence of a
 * generic converter contract: the contract names a code, not a code list, and
 * flattening an unrecognised one into `conversion_failed` would state a
 * vocabulary this server does not own -- and would erase the only string an
 * operator could search the proxy's own source for.
 */
const CONTRACT_ERROR_CODES = new Set([
  'unsupported_format',
  'too_large',
  'password_required',
  'conversion_failed',
  'fetch_failed',
  'proxy_unreachable',
  'invalid_target',
  'no_capacity',
]);

/**
 * Best-effort `name`/`contentType`/`size` for the error envelope.
 *
 * Probed only for mail and event attachments, and only on a failure path. Those
 * are the only Graph resources carrying all three fields; asking a message or a
 * photo for `$select=name,contentType,size` is a guaranteed 400, which would put
 * a noisy Graph error in the log on every failure and buy nothing.
 *
 * Every failure here is swallowed. The caller is already holding a real error,
 * and a probe that fails must never replace it -- "could not read the metadata
 * of the document you could not read" is strictly less useful than the original
 * refusal with three nulls beside it.
 */
async function describeAttachment(
  target: string,
  ctx: UtilityToolContext,
  accessToken: string | undefined
): Promise<AttachmentFacts> {
  if (!MAIL_EVENT_ATTACHMENT_TARGET.test(target) || !target.endsWith('/$value')) {
    return UNKNOWN_ATTACHMENT;
  }
  try {
    const metadataPath = target.slice(0, -'/$value'.length);
    const meta = (await ctx.graphClient.makeRequest(
      `${metadataPath}?$select=name,contentType,size`,
      { accessToken }
    )) as Record<string, unknown> | null;
    if (!meta || typeof meta !== 'object') return UNKNOWN_ATTACHMENT;
    return {
      name: typeof meta.name === 'string' ? meta.name : null,
      contentType: typeof meta.contentType === 'string' ? meta.contentType : null,
      size: typeof meta.size === 'number' ? meta.size : null,
    };
  } catch {
    return UNKNOWN_ATTACHMENT;
  }
}

export const UTILITY_TOOLS: readonly UtilityTool[] = [
  {
    name: 'parse-teams-url',
    method: 'POST',
    path: 'tool:parse-teams-url',
    description:
      'Converts any Teams meeting URL format (short /meet/, full /meetup-join/, or recap ?threadId=) into a standard joinWebUrl. Use this before list-online-meetings when the user provides a recap or short URL.',
    readOnlyHint: true,
    openWorldHint: false,
    buildSchema: () => ({
      url: z.string().describe('Teams meeting URL in any format'),
    }),
    execute: async (params) => {
      const url = params.url;
      if (typeof url !== 'string') {
        return {
          content: [{ type: 'text', text: JSON.stringify({ error: 'url is required.' }) }],
          isError: true,
        };
      }
      try {
        const joinWebUrl = parseTeamsUrl(url);
        return { content: [{ type: 'text', text: joinWebUrl }] };
      } catch (error) {
        return {
          content: [{ type: 'text', text: JSON.stringify({ error: (error as Error).message }) }],
          isError: true,
        };
      }
    },
  },
  {
    name: 'download-bytes',
    method: 'GET',
    path: 'tool:download-bytes',
    description:
      'Download binary content from Microsoft Graph and return it as base64. Single tool for any binary read: drive file content, mail attachment, profile photo, Teams hosted content, meeting recording. Returns { contentType, encoding: "base64", contentLength, contentBytes }. For large drive/SharePoint file content, prefer get-download-url, which returns a pre-authenticated URL to stream bytes out-of-band instead of base64 through the agent context. That preference always holds for drive/SharePoint files; for mail and event attachments, meeting recordings, and other /$value byte endpoints, get-download-url can only return a URL when the server runs with --enable-attachment-urls, so use this tool when it refuses.',
    readOnlyHint: true,
    openWorldHint: true,
    bytesToModel: true,
    buildSchema: (ctx) => {
      const schema: Record<string, z.ZodTypeAny> = {
        target: z
          .string()
          .describe(
            'Relative Microsoft Graph path starting with "/". Common paths: ' +
              '/drives/{drive-id}/items/{driveItem-id}/content (drive file content); ' +
              '/me/messages/{message-id}/attachments/{attachment-id}/$value (mail attachment, list-mail-attachments returns the IDs); ' +
              '/me/photo/$value or /users/{user-id}/photo/$value (profile photo); ' +
              '/chats/{chat-id}/messages/{chatMessage-id}/hostedContents/{chatMessageHostedContent-id}/$value (Teams chat hosted content, list-chat-message-hosted-contents returns the IDs); ' +
              '/teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id}/hostedContents/{chatMessageHostedContent-id}/$value (Teams channel hosted content). ' +
              'For meeting recordings, use get-meeting-recording-content where available; Microsoft Graph returns authenticated recording bytes, not a pre-authenticated download URL.'
          ),
      };
      if (ctx.multiAccount) {
        schema['account'] = z
          .string()
          .optional()
          .describe(
            'Account to use when multiple Microsoft accounts are configured. Required when multiple accounts exist (see list-accounts).'
          );
      }
      return schema;
    },
    execute: async (params, { graphClient, authManager }) => {
      const target = params.target;
      const accountParam = params.account as string | undefined;
      if (typeof target !== 'string' || target.length === 0) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: 'target is required and must be a non-empty string.' }),
            },
          ],
          isError: true,
        };
      }
      if (!target.startsWith('/')) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error:
                  'target must be a relative Microsoft Graph path starting with "/", e.g. /me/photo/$value or /drives/{drive-id}/items/{driveItem-id}/content. Absolute URLs are not accepted; if you have an @microsoft.graph.downloadUrl, use the equivalent /content or /$value path instead (Graph 302-redirects to the same bytes).',
              }),
            },
          ],
          isError: true,
        };
      }
      try {
        const accountModeError = await checkAccountParamInBearerMode(accountParam, authManager);
        if (accountModeError) {
          return {
            content: [{ type: 'text', text: JSON.stringify({ error: accountModeError }) }],
            isError: true,
          };
        }
        let accountAccessToken: string | undefined;
        if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
          accountAccessToken = await authManager.getTokenForAccount(accountParam);
        }
        // rawResponse keeps the body byte-faithful: binary stays base64 and a
        // JSON body is returned verbatim instead of being re-serialized lossily
        // through JSON.parse -> JSON.stringify (issue #546).
        return await graphClient.graphRequest(target, {
          accessToken: accountAccessToken,
          rawResponse: true,
        });
      } catch (error) {
        return {
          content: [{ type: 'text', text: JSON.stringify({ error: (error as Error).message }) }],
          isError: true,
        };
      }
    },
  },
  {
    name: 'download-bytes-to-file',
    method: 'GET',
    path: 'tool:download-bytes-to-file',
    searchKeywords:
      'save to disk save to file write file to disk save attachment to disk save recording to disk write bytes to local file output path',
    // Front-loaded on purpose: the discovery search index caps a tool's
    // description at ~40 tokens, so the OneDrive/SharePoint guidance below
    // sits past the cap. That keeps the hint for the reading LLM while letting
    // get-download-url own the high-signal "drive"/"sharepoint" search terms.
    description:
      'Write authenticated Microsoft Graph byte content to a local file on the server, returning { path, contentType, bytesWritten } instead of base64. The only out-of-band way to save mail attachments and meeting recordings, whose bytes are exposed solely through authenticated endpoints. Also handles profile photos and Teams hosted content. Writes to an absolute outputPath and never overwrites an existing file. stdio mode only: not available over HTTP. For OneDrive or SharePoint file content, get-download-url is preferred — it returns a pre-authenticated URL for fully out-of-band download without the server fetching the bytes.',
    readOnlyHint: true,
    openWorldHint: true,
    stdioOnly: true,
    buildSchema: (ctx) => {
      const schema: Record<string, z.ZodTypeAny> = {
        target: z
          .string()
          .describe(
            'Relative Microsoft Graph path starting with "/". Common paths: ' +
              '/drives/{drive-id}/items/{driveItem-id}/content (drive file content); ' +
              '/me/messages/{message-id}/attachments/{attachment-id}/$value (mail attachment, list-mail-attachments returns the IDs); ' +
              '/me/photo/$value or /users/{user-id}/photo/$value (profile photo); ' +
              '/chats/{chat-id}/messages/{chatMessage-id}/hostedContents/{chatMessageHostedContent-id}/$value (Teams chat hosted content, list-chat-message-hosted-contents returns the IDs); ' +
              '/teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id}/hostedContents/{chatMessageHostedContent-id}/$value (Teams channel hosted content). ' +
              'For meeting recordings, use get-meeting-recording-content where available; Microsoft Graph returns authenticated recording bytes, not a pre-authenticated download URL.'
          ),
        outputPath: z
          .string()
          .describe(
            "Absolute path on the server's filesystem where the bytes are written, e.g. /Users/me/downloads/invoice.pdf. Must be absolute; relative paths are rejected. The parent directory must already exist, and an existing file is never overwritten (the call errors if outputPath already exists)."
          ),
      };
      if (ctx.multiAccount) {
        schema['account'] = z
          .string()
          .optional()
          .describe(
            'Account to use when multiple Microsoft accounts are configured. Required when multiple accounts exist (see list-accounts).'
          );
      }
      return schema;
    },
    execute: async (params, { graphClient, authManager }) => {
      const target = params.target;
      const outputPath = params.outputPath;
      const accountParam = params.account as string | undefined;
      if (typeof target !== 'string' || target.length === 0) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: 'target is required and must be a non-empty string.' }),
            },
          ],
          isError: true,
        };
      }
      if (!target.startsWith('/')) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error:
                  'target must be a relative Microsoft Graph path starting with "/", e.g. /me/photo/$value or /drives/{drive-id}/items/{driveItem-id}/content. Absolute URLs are not accepted; if you have an @microsoft.graph.downloadUrl, use the equivalent /content or /$value path instead (Graph 302-redirects to the same bytes).',
              }),
            },
          ],
          isError: true,
        };
      }
      if (typeof outputPath !== 'string' || outputPath.length === 0) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: 'outputPath is required and must be a non-empty string.',
              }),
            },
          ],
          isError: true,
        };
      }
      if (!path.isAbsolute(outputPath)) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: `outputPath must be an absolute path, e.g. /Users/me/downloads/file.ext. Received: ${outputPath}`,
              }),
            },
          ],
          isError: true,
        };
      }
      // downloadToFile's wx is the real no-overwrite guard; this just gives a
      // friendlier "already exists" error before we bother calling Graph.
      let fileExists = false;
      try {
        await access(outputPath);
        fileExists = true;
      } catch {
        // ENOENT (and any other access error) means the file isn't readable/there;
        // let the write attempt surface the real problem.
      }
      if (fileExists) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: `file already exists at ${outputPath}` }),
            },
          ],
          isError: true,
        };
      }
      try {
        const accountModeError = await checkAccountParamInBearerMode(accountParam, authManager);
        if (accountModeError) {
          return {
            content: [{ type: 'text', text: JSON.stringify({ error: accountModeError }) }],
            isError: true,
          };
        }
        let accountAccessToken: string | undefined;
        if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
          accountAccessToken = await authManager.getTokenForAccount(accountParam);
        }
        // Stream to disk instead of buffering: makeRequest holds the whole file
        // in memory as base64, which dies on big recordings (V8 max string length).
        const result = await graphClient.downloadToFile(target, outputPath, {
          accessToken: accountAccessToken,
        });
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                path: outputPath,
                contentType: result.contentType,
                bytesWritten: result.contentLength,
              }),
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: 'text', text: JSON.stringify({ error: (error as Error).message }) }],
          isError: true,
        };
      }
    },
  },
  {
    name: 'get-download-url',
    method: 'GET',
    path: 'tool:get-download-url',
    searchKeywords:
      'download file download drive file download onedrive file sharepoint file download large drive file large sharepoint file large file out-of-band download pre-authenticated url',
    // Front-loaded on purpose, same constraint as download-bytes-to-file: the
    // discovery search index caps a tool's description at ~40 tokens, so
    // "drive/SharePoint file content" has to stay in the opening sentence or this
    // tool stops owning the "drive"/"sharepoint" download queries.
    description:
      "Resolve a short-lived, pre-authenticated download URL for Microsoft Graph binary content: always available for drive/SharePoint file content, and for other byte endpoints only when this server was started with --enable-attachment-urls. The returned URL streams the bytes with NO Authorization header, so the client can fetch it straight to disk (e.g. curl) without round-tripping base64 through the agent context. Prefer this over download-bytes for any file above a few KB or any bulk download. For a drive/SharePoint item the URL is Graph's own @microsoft.graph.downloadUrl and needs no flag; returns { downloadUrl, name?, size?, contentType? }. Mail and event attachments (/messages/{id}/attachments/{id}/$value), meeting recordings, and other authenticated /$value byte endpoints have no such link from Graph, so with --enable-attachment-urls (HTTP mode only) this server mints and serves one itself, returning { downloadUrl, expiresAt, maxFetches, singleUse: false, note, onFetchFailure } — good for up to " +
      `${MAX_REDEMPTIONS} fetches until expiresAt, NOT one. Hand the same URL to a document converter more than once: probing a document and then converting it works, as does a pagination continuation. Every fetch counts, a failed one included, so when a fetch fails retry that same URL rather than minting another; only a 404 means it is finished (fetches used up, or expired) and only then mint again. Without the flag those targets fail with an error saying they do not expose a pre-authenticated download URL; fall back to download-bytes. Minting is also refused whenever this request's Graph identity came from the caller rather than from the server's own token cache (OAuth, OBO, or bearer mode), because the minted URL is redeemed later with no Authorization header and would fetch the bytes under a different identity than the one that asked; in those modes use download-bytes.`,
    readOnlyHint: true,
    openWorldHint: true,
    bytesToModel: true,
    buildSchema: (ctx) => {
      const schema: Record<string, z.ZodTypeAny> = {
        target: z
          .string()
          .describe(
            'Relative Microsoft Graph path starting with "/". Either a driveItem content path or the item path itself, e.g. ' +
              '/drives/{drive-id}/items/{driveItem-id}/content, /me/drive/items/{driveItem-id}/content, ' +
              'or /sites/{site-id}/drive/items/{driveItem-id}. ' +
              'A trailing /content is optional and is stripped automatically for drive items. Mail and event attachment $value paths, meeting recordings, and other /$value byte endpoints are accepted only when the server runs with --enable-attachment-urls (Graph exposes no pre-authenticated URL for them, so the server mints one); otherwise they are rejected and download-bytes is the fallback.'
          ),
      };
      if (ctx.multiAccount) {
        schema['account'] = z
          .string()
          .optional()
          .describe(
            'Account to use when multiple Microsoft accounts are configured. Required when multiple accounts exist (see list-accounts).'
          );
      }
      return schema;
    },
    execute: async (params, { graphClient, authManager }) => {
      const target = params.target;
      const accountParam = params.account as string | undefined;
      if (typeof target !== 'string' || target.length === 0) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: 'target is required and must be a non-empty string.' }),
            },
          ],
          isError: true,
        };
      }
      if (!target.startsWith('/')) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error:
                  'target must be a relative Microsoft Graph path starting with "/", e.g. /drives/{drive-id}/items/{driveItem-id}/content.',
              }),
            },
          ],
          isError: true,
        };
      }
      // Normalize: separate any query string and strip trailing slashes so the /content and
      // /$value suffix checks are robust to e.g. "/content/" or "/content?select=id".
      const queryIdx = target.indexOf('?');
      if (queryIdx >= 0) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error:
                  'target must not include query parameters. Pass the drive item /content path or item metadata path without $select, $expand, or other query options.',
              }),
            },
          ],
          isError: true,
        };
      }
      const pathPart = target.replace(/\/+$/, '');
      // Mail/event attachments expose no pre-authenticated download URL in Graph; bytes come
      // only from base64 contentBytes or the authenticated /$value endpoint (use download-bytes).
      // Match only real Graph mail/calendar attachment resources so driveItem path addressing
      // with folders named messages/events/attachments is not falsely rejected.
      if (MAIL_EVENT_ATTACHMENT_TARGET.test(pathPart)) {
        const minted = await mintDownloadUrl(pathPart, accountParam, authManager);
        if (minted) return minted;
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error:
                  'Mail and calendar event attachments do not expose a pre-authenticated download URL. Use download-bytes for small attachments.',
              }),
            },
          ],
          isError: true,
        };
      }
      // Recording content endpoints return authenticated bytes, not a pre-authenticated URL.
      if (MEETING_RECORDING_TARGETS.some((pattern) => pattern.test(pathPart))) {
        const minted = await mintDownloadUrl(pathPart, accountParam, authManager);
        if (minted) return minted;
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error:
                  'Meeting recordings do not expose a pre-authenticated download URL. Use download-bytes for small recordings or get-meeting-recording-content where available.',
              }),
            },
          ],
          isError: true,
        };
      }
      // Other /$value byte endpoints (profile photo, Teams hosted content) likewise have no URL.
      if (VALUE_BYTE_TARGET.test(pathPart)) {
        const minted = await mintDownloadUrl(pathPart, accountParam, authManager);
        if (minted) return minted;
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error:
                  '$value byte endpoints do not expose a pre-authenticated download URL. Use download-bytes to read these bytes.',
              }),
            },
          ],
          isError: true,
        };
      }
      const isDriveItemById =
        /^\/drives\/[^/]+\/items\/[^/]+(?:\/content)?$/.test(pathPart) ||
        /^\/(?:me|users\/[^/]+|groups\/[^/]+|sites\/[^/]+)\/drive\/items\/[^/]+(?:\/content)?$/.test(
          pathPart
        ) ||
        /^\/(?:groups\/[^/]+|sites\/[^/]+)\/drives\/[^/]+\/items\/[^/]+(?:\/content)?$/.test(
          pathPart
        );
      const isDriveItemByPath =
        /^\/drives\/[^/]+\/root:\/.+:(?:\/content)?$/.test(pathPart) ||
        /^\/(?:me|users\/[^/]+|groups\/[^/]+|sites\/[^/]+)\/drive\/root:\/.+:(?:\/content)?$/.test(
          pathPart
        ) ||
        /^\/(?:groups\/[^/]+|sites\/[^/]+)\/drives\/[^/]+\/root:\/.+:(?:\/content)?$/.test(
          pathPart
        );
      if (!isDriveItemById && !isDriveItemByPath) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error:
                  'target must identify a driveItem in OneDrive or SharePoint. Use a drive item metadata path or /content path, such as /drives/{drive-id}/items/{driveItem-id}/content, /me/drive/items/{driveItem-id}, /sites/{site-id}/drive/items/{driveItem-id}, or /me/drive/root:/path/file.ext:/content. Other Graph byte resources must use download-bytes.',
              }),
            },
          ],
          isError: true,
        };
      }
      // The downloadUrl lives on driveItem metadata, not the /content sub-resource.
      // Only strip true Graph content endpoints: ID-addressed /items/{id}/content
      // and path-addressed root:/path/file:/content. A drive item can itself be
      // named "content", so a plain trailing /content is not enough.
      const isDriveContentEndpoint =
        /\/items\/[^/]+\/content$/.test(pathPart) || pathPart.endsWith(':/content');
      const itemPath = isDriveContentEndpoint ? pathPart.slice(0, -'/content'.length) : pathPart;
      try {
        const accountModeError = await checkAccountParamInBearerMode(accountParam, authManager);
        if (accountModeError) {
          return {
            content: [{ type: 'text', text: JSON.stringify({ error: accountModeError }) }],
            isError: true,
          };
        }

        let accountAccessToken: string | undefined;
        if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
          accountAccessToken = await authManager.getTokenForAccount(accountParam);
        }
        const response = await graphClient.graphRequest(itemPath, {
          accessToken: accountAccessToken,
          // We JSON.parse the metadata below, so force JSON - under --toon it'd be
          // TOON and the parse would fail, masking a real item as "no download url".
          forceJsonOutput: true,
        });
        // graphRequest swallows Graph HTTP errors and returns { isError: true } (see
        // graph-client.ts); surface the real error (401/403/404/429/...) instead of masking
        // it as "no download URL available".
        if (response?.isError) {
          return response;
        }
        const text = response?.content?.[0]?.text;
        let item: Record<string, unknown> | undefined;
        if (typeof text === 'string') {
          try {
            item = JSON.parse(text);
          } catch {
            item = undefined;
          }
        }
        const downloadUrl = item?.['@microsoft.graph.downloadUrl'] as string | undefined;
        if (!downloadUrl) {
          // The metadata path carries no downloadUrl, but the bytes are still
          // reachable at its /content sub-resource -- which is what a ticket
          // has to name, since that is what the redemption route will GET.
          const minted = await mintDownloadUrl(`${itemPath}/content`, accountParam, authManager);
          if (minted) return minted;
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  error:
                    'No pre-authenticated download URL is available for this resource. It may not be a drive item, or it exposes bytes only via download-bytes.',
                }),
              },
            ],
            isError: true,
          };
        }
        const file = item?.file as { mimeType?: string } | undefined;
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                downloadUrl,
                name: item?.name,
                size: item?.size,
                contentType: file?.mimeType,
              }),
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: 'text', text: JSON.stringify({ error: (error as Error).message }) }],
          isError: true,
        };
      }
    },
  },
  {
    name: 'read-document',
    method: 'POST',
    path: 'tool:read-document',
    searchKeywords:
      'read attachment read document convert to markdown pdf docx xlsx pptx eml msg extract text from attachment open attachment',
    description:
      'Read any Microsoft 365 document as markdown: mail and event attachments, OneDrive and SharePoint files, and raw message MIME. Give it the Graph byte path (list-mail-attachments returns the ids) and it returns text, never bytes — this server fetches the document itself and converts it out of band, so nothing base64 ever enters this conversation. Supports paging via pages/offset/maxChars for long documents. This is the ONLY way to read the content of anything inside Microsoft 365 on this server. For anything OUTSIDE Microsoft 365 — a public web URL, a link found in an email body — use the document converter tool directly instead.',
    readOnlyHint: true,
    openWorldHint: true,
    proxyOnly: true,
    buildSchema: (ctx) => {
      const schema: Record<string, z.ZodTypeAny> = {
        target: z
          .string()
          .describe(
            'Relative Microsoft Graph byte path starting with "/". ' +
              '/me/messages/{message-id}/attachments/{attachment-id}/$value (mail attachment; list-mail-attachments returns the ids); ' +
              '/me/events/{event-id}/attachments/{attachment-id}/$value (event attachment); ' +
              '/me/messages/{message-id}/$value (the whole message as RFC 5322 source); ' +
              '/drives/{drive-id}/items/{driveItem-id}/content (drive or SharePoint file). ' +
              'Absolute URLs are not accepted.'
          ),
        pages: z
          .string()
          .optional()
          .describe(
            'Page selection for paged formats, e.g. "1-5" or "2,4,9". Omit for the whole document.'
          ),
        offset: z
          .number()
          .optional()
          .describe('Character offset to resume from, for continuing a long read.'),
        maxChars: z
          .number()
          .optional()
          .describe('Maximum characters of markdown to return in this call.'),
      };
      if (ctx.multiAccount) {
        schema['account'] = z
          .string()
          .optional()
          .describe(
            'Account to use when multiple Microsoft accounts are configured. Required when multiple accounts exist (see list-accounts).'
          );
      }
      return schema;
    },
    execute: async (params, ctx) => {
      const target = params.target;
      if (typeof target !== 'string' || target.length === 0) {
        return readDocumentError(
          'invalid_target',
          'target is required and must be a non-empty relative Microsoft Graph path starting with "/".',
          UNKNOWN_ATTACHMENT
        );
      }

      const proxy = getAttachmentProxy();
      const minting = getAttachmentMinting();
      if (!proxy || !minting) {
        return readDocumentError(
          'proxy_unreachable',
          'This server has no document proxy configured, so no document can be read. It must be started with --attachment-proxy and --http.',
          UNKNOWN_ATTACHMENT
        );
      }

      // Validated against the same patterns get-download-url mints for, and for
      // the same reason: a ticket grants an authenticated GET of exactly one
      // Graph path with this server's own token, so the set of paths a ticket
      // can name is the whole of what the capability is worth. One list, one
      // answer to "what can this feature reach".
      if (!target.startsWith('/') || !MINTABLE_TARGET_PATTERNS.some((p) => p.test(target))) {
        return readDocumentError(
          'invalid_target',
          `target must be a relative Microsoft Graph byte path this server can mint for: a mail or event attachment ` +
            `(/me/messages/{message-id}/attachments/{attachment-id}/$value), a meeting recording, or any other ` +
            `authenticated /$value endpoint (/me/messages/{message-id}/$value for the raw message). Absolute URLs ` +
            `are not accepted. Got ${JSON.stringify(target)}.`,
          UNKNOWN_ATTACHMENT
        );
      }

      const accountParam = params.account as string | undefined;

      // The identity guard get-download-url already carries, restated because
      // read-document mints too. Both halves matter: isOAuthModeEnabled() is
      // false in plain bearer mode and in --obo, both of which still run inside
      // a request context holding the CALLER's token, while a redeemed ticket
      // is fetched with the SERVER's. Minting there would let a caller ask under
      // one identity and have the bytes read under another.
      if (ctx.authManager?.isOAuthModeEnabled() || getRequestTokens()) {
        return readDocumentError(
          'identity_not_supported',
          'read-document is unavailable when Graph identity comes from the request (OAuth, OBO, or bearer mode): the minted URL is redeemed later with no Authorization header, so the document would be fetched as a different identity than the one that asked for it.',
          UNKNOWN_ATTACHMENT
        );
      }

      const accountModeError = await checkAccountParamInBearerMode(accountParam, ctx.authManager);
      if (accountModeError) {
        return readDocumentError('identity_not_supported', accountModeError, UNKNOWN_ATTACHMENT);
      }

      // The server's own token, resolved once. In proxy mode identity always
      // comes from the token cache -- the guard above refused every other mode --
      // so this is the same identity the redemption route will use.
      let accessToken: string | undefined;
      try {
        accessToken = await ctx.authManager?.getTokenForAccount(accountParam);
      } catch {
        // Left undefined: makeRequest resolves its own token, and a token
        // problem will surface as the Graph error it is rather than here.
        accessToken = undefined;
      }

      /**
       * One attempt: one FRESH mint, one conversion.
       *
       * Fresh per attempt, not per call. A failed fetch still spends a
       * redemption and the proxy may have spent one or more before failing, so
       * a retry on the same ticket can meet a 404 that has nothing to do with
       * why the first attempt failed. Minting server-side is what makes this
       * affordable: the 3-redemption budget and the 120 s TTL stopped being
       * agent-visible the moment the agent stopped holding the URL.
       *
       * Redaction happens here, per attempt, against THIS attempt's own
       * ticket id and URL -- not once at the end against whichever ticket
       * happened to be minted last. A retry mints a second, different ticket,
       * so redacting the final outcome with only the second ticket's id would
       * leave the first ticket's id exposed in a message the first attempt
       * produced (moot today, since a proxy_unreachable message never carries
       * a ticket, but this must hold for every future code, not only the ones
       * observed so far).
       */
      // Exactly one retry: two attempts, numbered for the warn line below.
      const MAX_ATTEMPTS = 2;

      const attempt = async (
        attemptNumber: number
      ): Promise<{ ok: true; markdown: string } | { ok: false; code: string; message: string }> => {
        let ticket: { id: string; expiresAtMs: number };
        try {
          ticket = minting.store.mint(target, accountParam);
        } catch (error) {
          if (error instanceof TicketStoreFullError) {
            return { ok: false, code: 'no_capacity', message: error.message };
          }
          throw error;
        }
        const ticketUrl = buildAttachmentUrl(minting.config, ticket.id);
        const startedAtMs = Date.now();
        const outcome = await proxy.client.convertToMarkdown({
          uri: ticketUrl,
          ...(typeof params.pages === 'string' ? { pages: params.pages } : {}),
          ...(typeof params.offset === 'number' ? { offset: params.offset } : {}),
          ...(typeof params.maxChars === 'number' ? { maxChars: params.maxChars } : {}),
        });
        if (!outcome.ok && outcome.code === 'proxy_unreachable') {
          // The ONE warn for this failure, deliberately not duplicated by the
          // client (`AttachmentProxyClient` logs the same condition at debug,
          // not warn -- see attachment-proxy.ts). Only this layer knows the
          // attempt number, and "attempt 1 of 2" versus "still unreachable
          // after retry" is exactly what tells an operator a blip from a
          // wedge in `docker logs m365-max-mcp`. The proxy ships a liveness
          // healthcheck that never consults its own workers, so a wedged one
          // still reports healthy -- exactly the shape of the 19-hour silent
          // failure this stack has already seen, and a signal worth keeping
          // singular and unambiguous rather than doubling it across layers.
          const attemptNote =
            attemptNumber >= MAX_ATTEMPTS
              ? `attempt ${attemptNumber} of ${MAX_ATTEMPTS}, still unreachable after retry`
              : `attempt ${attemptNumber} of ${MAX_ATTEMPTS}`;
          logger.warn(
            `Attachment proxy unreachable (${attemptNote}): ${proxy.url} did not answer after ` +
              `${Date.now() - startedAtMs}ms (${outcome.message})`
          );
        }
        // Redacted on both branches: the proxy was handed the live ticket URL as
        // its `uri` argument, and nothing stops it from echoing that URL (or
        // just the ticket id) back inside EITHER a converted document's content
        // or an error message. `outcome.code` is included too, defensively --
        // today it is always one of CONTRACT_ERROR_CODES or the fixed literal
        // 'proxy_error' (see attachment-proxy.ts's mapProxyError /
        // interpretJsonRpcMessage), never proxy-chosen free text, so this redact
        // is a no-op on the current contract rather than a gap it is closing.
        if (outcome.ok) {
          return {
            ok: true,
            markdown: redactAttachmentSecrets(outcome.markdown, ticket.id, ticketUrl),
          };
        }
        return {
          ok: false,
          code: redactAttachmentSecrets(outcome.code, ticket.id, ticketUrl),
          message: redactAttachmentSecrets(outcome.message, ticket.id, ticketUrl),
        };
      };

      let outcome = await attempt(1);
      // Exactly one retry, and only for the transport class. A proxy that
      // ANSWERED (too_large, password_required, ...) will answer the same way
      // again, so retrying would double the conversion cost for no new
      // information; a connection that never landed might.
      if (!outcome.ok && outcome.code === 'proxy_unreachable') {
        outcome = await attempt(2);
      }

      if (outcome.ok) {
        return { content: [{ type: 'text', text: outcome.markdown }] };
      }

      const facts = await describeAttachment(target, ctx, accessToken);
      const known = CONTRACT_ERROR_CODES.has(outcome.code);
      return readDocumentError(
        known ? outcome.code : 'proxy_error',
        outcome.message,
        facts,
        known ? undefined : outcome.code
      );
    },
  },
];

/**
 * Is this a GET whose Graph path ends in `/$value`?
 *
 * Narrower than "does this tool return raw bytes to the model" -- deliberately
 * so; read that broader claim off this function at your peril. `/$value` is
 * Graph's own spelling for "the raw representation of this resource", so a GET
 * ending there returns bytes by construction, and a class rule over that shape
 * (rather than a name list, which is the same story-shaped guard one level
 * down) covers a future endpoint upstream adds with no edit here. Today it
 * selects exactly `get-mail-message-mime`; the only other `$value` endpoint in
 * endpoints.json is a PUT (upload-my-profile-photo), which writes bytes rather
 * than returning them and is correctly left alone.
 *
 * It does NOT cover every raw-byte read this server exposes, and callers must
 * not treat "suppressed by this rule" as "the only bytes left." Known gaps,
 * left open for a scoped follow-up rather than widened here without its own
 * review:
 *  - `graph-batch` (POST `/$batch`) accepts arbitrary sub-requests and can
 *    smuggle a GET against any suppressed `/$value` path -- including
 *    `/me/messages/{id}/$value` -- as a batched sub-request, returning the
 *    same bytes this function exists to keep out. Suppressing a
 *    general-purpose batch tool is a capability decision, not a class-rule fix.
 *  - `get-meeting-recording-content` (video), `get-meeting-transcript-content`
 *    (text/vtt), `get-onenote-page-content`, and
 *    `get-sharepoint-site-onenote-page-content` are raw-byte/text reads whose
 *    paths do not end in `/$value`, so this rule does not see them.
 *
 * get-mail-message-mime genuinely cannot be left to the response scrubber: it
 * declares `acceptType: "text/plain"` and returns RFC 5322 source, not base64,
 * so neither scrubber rule matches, while every attachment rides inline. The
 * same is true of a batched read of the same path via `graph-batch`.
 */
export function isProxySuppressedGraphTool(
  method: string,
  pathPattern: string | undefined
): boolean {
  return method.toUpperCase() === 'GET' && /\/\$value$/.test(pathPattern ?? '');
}

/** Every gate that can keep a utility tool out of the registered set. */
export interface UtilityToolGates {
  readOnly?: boolean;
  httpMode?: boolean;
  /** Raw --enabled-tools / --preset pattern. An uncompilable pattern is ignored, as at registration. */
  enabledTools?: string;
  /** --attachment-proxy: byte-returning tools out, read-document in. */
  attachmentProxy?: boolean;
}

function compileToolFilter(pattern?: string): RegExp | undefined {
  if (!pattern) return undefined;
  try {
    return new RegExp(pattern, 'i');
  } catch {
    // Registration logs and then ignores an invalid pattern, exposing everything. Mirror that
    // here rather than guessing, so a predicate built from these gates can never disagree with
    // what actually gets registered.
    return undefined;
  }
}

/**
 * The single definition of which utility tools a given configuration registers.
 *
 * Both registration paths (registerGraphTools and registerDiscoveryTools) select through this, and
 * so does the startup check that warns when a flag has been enabled but the only tool that can act
 * on it was filtered away. Duplicating these three conditions is how "enabled, validated, and
 * unreachable" happens in the first place: the gate and the warning drift apart and the warning
 * stops describing the server.
 */
export function selectUtilityTools(gates: UtilityToolGates): UtilityTool[] {
  const enabledToolsRegex = compileToolFilter(gates.enabledTools);
  return UTILITY_TOOLS.filter((utility) => {
    if (gates.readOnly && !utility.readOnlyHint) return false;
    if (gates.httpMode && utility.stdioOnly) return false;
    if (gates.attachmentProxy && utility.bytesToModel) return false;
    if (!gates.attachmentProxy && utility.proxyOnly) return false;
    if (enabledToolsRegex && !enabledToolsRegex.test(utility.name)) return false;
    return true;
  });
}

export function utilityToolWillRegister(name: string, gates: UtilityToolGates): boolean {
  return selectUtilityTools(gates).some((utility) => utility.name === name);
}

function registerUtilityToolWithMcp(
  server: McpServer,
  utility: UtilityTool,
  ctx: UtilityToolContext
): void {
  server.tool(
    utility.name,
    utility.description,
    utility.buildSchema(ctx),
    {
      title: utility.name,
      readOnlyHint: utility.readOnlyHint ?? true,
      openWorldHint: utility.openWorldHint ?? true,
    },
    async (params) => utility.execute(params, ctx)
  );
}

// Dig out the object shape of a Body schema so flattened top-level params can be
// matched against it (#569). z.lazy (chatMessage etc.) hides it behind _def.getter
function bodySchemaShape(schema: z.ZodTypeAny | undefined): Record<string, unknown> | null {
  let current: z.ZodTypeAny | undefined = schema;
  for (let i = 0; i < 10 && current; i++) {
    if (current instanceof z.ZodObject) {
      return current.shape as Record<string, unknown>;
    }
    const def = (
      current as {
        _def?: { innerType?: z.ZodTypeAny; schema?: z.ZodTypeAny; getter?: () => z.ZodTypeAny };
      }
    )._def;
    current = def?.innerType ?? def?.schema ?? def?.getter?.();
  }
  return null;
}

// SDK validation hands the handler the PARSED value, and strip-mode objects silently
// drop unknown keys - passthrough keeps whatever the client sent
function lenientBodySchema(schema: z.ZodTypeAny): z.ZodTypeAny {
  if (schema instanceof z.ZodObject) {
    return schema.passthrough();
  }
  if (schema instanceof z.ZodOptional) {
    return lenientBodySchema(schema.unwrap()).optional();
  }
  if (schema instanceof z.ZodNullable) {
    return lenientBodySchema(schema.unwrap()).nullable();
  }
  if (schema instanceof z.ZodLazy) {
    return z.lazy(() => lenientBodySchema(schema.schema));
  }
  return schema;
}

// Read-only in Graph - merging an echoed id/timestamp into a POST/PATCH body can 400
const READ_ONLY_BODY_FIELDS = new Set([
  'id',
  'createdDateTime',
  'lastModifiedDateTime',
  'changeKey',
]);

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

// Object.hasOwn, but tsconfig targets ES2020. Not `in` - that would match
// toString/constructor through the prototype
function hasOwn(obj: Record<string, unknown>, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(obj, key);
}

/**
 * Navigation properties whose expansion inlines base64 content into the parent
 * resource.
 *
 * Property NAMES, not (tool, property) pairs. The same navigation property is
 * expandable from every tool that reaches the same entity -- get-mail-message,
 * list-mail-messages, the delta tools, execute-tool -- so a pair list would have
 * to be re-derived every time an endpoint is added, and would be wrong the first
 * time one was missed.
 *
 * Known gap, left open for a scoped follow-up rather than widened here without
 * its own review: `graph-batch` (POST `/$batch`) accepts arbitrary sub-request
 * URLs, e.g. `{ url: "/me/messages/{id}?$expand=attachments" }`, and this guard
 * only inspects the top-level `expand`/`$expand` parameters `findByteInliningExpand`
 * is handed -- it does not parse batch sub-request URLs, so a batch payload can
 * smuggle the same byte-inlining expand this guard exists to keep out.
 * Suppressing a general-purpose batch tool is a capability decision, not a
 * class-rule fix; see the identical gap disclosed on `isProxySuppressedGraphTool`
 * above.
 */
const BYTE_INLINING_NAV_PROPERTIES = new Set(['attachments', 'hostedcontents']);

/**
 * Every occurrence of `expand=` (with or without a leading `$`, any casing)
 * inside `text`, together with the value that follows it up to the matching
 * unbalanced `)` or the end of the string.
 *
 * OData nests a sub-resource's own query options inside `(...)` after the
 * navigation property, e.g. `instances($expand=attachments)` for a recurring
 * event's expanded instances -- a real Graph pattern, and the reason the
 * top-level head-token check in `findByteInliningExpand` alone is not enough:
 * its head is `instances`, so `attachments` living inside the parens is never
 * seen by a check that only looks before the first `(`.
 *
 * A manual balanced scan rather than a single regex, so that a captured value
 * which itself contains `(...)` (deeper nesting, e.g. a doubly-nested
 * `$expand`) does not get truncated at the first `)` -- that inner paren is
 * consumed as part of the value, and the scan only stops at the `)` that
 * closes the *enclosing* group. Combined with the recursive call in
 * `scanForByteInliningHead`, this is what makes detection depth-independent:
 * each extracted value is fed back through the same scan, which finds and
 * extracts any `expand=` nested inside it, and so on.
 */
function extractNestedExpandValues(text: string): string[] {
  const values: string[] = [];
  const pattern = /\$?expand\s*=\s*/gi;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(text)) !== null) {
    const start = match.index + match[0].length;
    let depth = 0;
    let end = start;
    while (end < text.length) {
      const ch = text[end];
      if (ch === '(') {
        depth++;
      } else if (ch === ')') {
        if (depth === 0) break;
        depth--;
      }
      end++;
    }
    values.push(text.slice(start, end));
  }
  return values;
}

/**
 * The offending token within one `expand` entry, or null.
 *
 * Handles every spelling a caller can produce: a comma-separated list inside
 * one string, an OData nested option suffix (`attachments($select=name)`), a
 * type-cast path segment (`attachments/microsoft.graph.fileAttachment`), any
 * casing, surrounding whitespace, and -- via `extractNestedExpandValues` --
 * a `$expand` nested inside another property's parenthesised options, at any
 * nesting depth.
 *
 * Splitting on `,` also splits inside a nested option list, which is fine for
 * detection: the head token of `attachments($select=id,name)` is always in the
 * first fragment, so a false negative cannot arise from the split.
 */
function scanForByteInliningHead(text: string): string | null {
  for (const piece of text.split(',')) {
    const trimmed = piece.trim();
    if (!trimmed) continue;
    const head = trimmed.split('(')[0].split('/')[0].trim().toLowerCase();
    if (BYTE_INLINING_NAV_PROPERTIES.has(head)) return trimmed;
  }
  for (const nested of extractNestedExpandValues(text)) {
    const found = scanForByteInliningHead(nested);
    if (found) return found;
  }
  return null;
}

/**
 * The offending `$expand` entry, or null.
 *
 * Reads `$expand` and `expand` independently -- never with `??` -- and scans
 * every entry found under EITHER key. `.passthrough()` on every tool's input
 * schema (and `execute-tool`'s `z.record(z.any())` parameters) means a caller
 * can hand this function both keys at once, e.g. `{ $expand: [], expand:
 * ['attachments'] }`; picking one key over the other would let a present but
 * empty `$expand` mask a harmful `expand`, or vice versa.
 */
export function findByteInliningExpand(params: Record<string, unknown>): string | null {
  const rawValues = [params.$expand, params.expand].filter(
    (raw) => raw !== undefined && raw !== null
  );
  for (const raw of rawValues) {
    const entries = Array.isArray(raw) ? raw : [raw];
    for (const entry of entries) {
      if (typeof entry !== 'string') continue;
      const found = scanForByteInliningHead(entry);
      if (found) return found;
    }
  }
  return null;
}

async function executeGraphTool(
  tool: (typeof api.endpoints)[0],
  config: EndpointConfig | undefined,
  graphClient: GraphClient,
  params: Record<string, unknown>,
  authManager?: AuthManager
): Promise<CallToolResult> {
  logger.info(`Tool ${tool.alias} called with params: ${JSON.stringify(params)}`);

  if (
    isConfirmGateEnabled() &&
    isDestructiveOperation(tool.method, config) &&
    params.confirm !== true
  ) {
    logger.warn(
      `Refusing destructive tool ${tool.alias} (${tool.method.toUpperCase()}): missing confirm: true`
    );
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            error: 'confirmation_required',
            tool: tool.alias,
            method: tool.method.toUpperCase(),
            destructive: true,
            message:
              'This tool modifies user data. Re-call with parameter "confirm": true after the user has explicitly approved the operation.',
          }),
        },
      ],
      isError: true,
    };
  }

  // Refused once, here, because this is where both paths land: the handler
  // registerGraphTools installs on every tool, and discovery's execute-tool.
  // Guarding the schema instead would cover 38 tools one at a time and miss the
  // 39th.
  if (getAttachmentProxy()) {
    const blocked = findByteInliningExpand(params);
    if (blocked) {
      logger.warn(
        `Refusing ${tool.alias}: expand "${blocked}" would inline attachment bytes (--attachment-proxy)`
      );
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              error: 'expand_not_allowed',
              tool: tool.alias,
              expand: blocked,
              message:
                `Expanding "${blocked}" inlines the raw attachment bytes (base64 contentBytes) into this ` +
                `response, and $select does not suppress them. This server runs with --attachment-proxy, where ` +
                `no tool returns raw bytes. Call this tool again WITHOUT that expand value to get the message ` +
                `or event itself; use list-mail-attachments (or the matching list tool) for each attachment's ` +
                `id, name, contentType and size; and use read-document with the attachment's $value path to ` +
                `read its content as markdown.`,
            }),
          },
        ],
        isError: true,
      };
    }
  }

  const requestId = randomUUID();
  const startTime = Date.now();
  const upn = getUserIdentityForAudit(getRequestTokens()?.accessToken);
  const httpMethod = tool.method.toUpperCase();
  let targetResource: AuditTargetResource | undefined;

  try {
    const accountParam = params.account as string | undefined;

    // In OAuth/HTTP bearer mode, refuse an `account` param that doesn't match the bearer
    // identity instead of silently returning the bearer user's data (discussion #467).
    const accountModeError = await checkAccountParamInBearerMode(accountParam, authManager);
    if (accountModeError) {
      return {
        content: [{ type: 'text', text: JSON.stringify({ error: accountModeError }) }],
        isError: true,
      };
    }

    // Resolve account-specific token if `account` parameter is provided (or auto-resolve for single account).
    // Skip in OAuth/HTTP mode — let the request context drive token selection via GraphClient.
    // Also skip when a request-context token exists (HTTP/OAuth flow where token comes from middleware).
    let accountAccessToken: string | undefined;
    if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
      try {
        accountAccessToken = await authManager.getTokenForAccount(accountParam);
      } catch (err) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: (err as Error).message }),
            },
          ],
          isError: true,
        };
      }
    }

    const parameterDefinitions = tool.parameters || [];

    let path = tool.path;
    const queryParams: Record<string, string> = {};
    const headers: Record<string, string> = {};
    let body: unknown = null;

    // Body fields the client passed as top-level params (#569) - merged into the
    // request body after the loop
    const bodyShape = bodySchemaShape(
      parameterDefinitions.find((p) => p.type === 'Body')?.schema as z.ZodTypeAny | undefined
    );
    const strayBodyFields: Record<string, unknown> = {};

    for (const [paramName, paramValue] of Object.entries(params)) {
      // Skip control parameters - not part of the Microsoft Graph API
      if (
        [
          'account',
          'confirm',
          'fetchAllPages',
          'includeHeaders',
          'excludeResponse',
          'timezone',
          'expandExtendedProperties',
        ].includes(paramName)
      ) {
        continue;
      }

      // Ok, so, MCP clients (such as claude code) doesn't support $ in parameter names,
      // and others might not support __, so we strip them in hack.ts and restore them here
      const odataParams = [
        'filter',
        'select',
        'expand',
        'orderby',
        'skip',
        'top',
        'count',
        'search',
        'format',
      ];
      // Handle both "top" and "$top" formats - strip $ if present, then re-add it
      const normalizedParamName = paramName.startsWith('$') ? paramName.slice(1) : paramName;
      const isOdataParam = odataParams.includes(normalizedParamName.toLowerCase());
      const fixedParamName = isOdataParam ? `$${normalizedParamName.toLowerCase()}` : paramName;
      // Convert kebab-case param names to camelCase for path param matching.
      // endpoints.json uses {message-id} but hack.ts extracts :messageId (camelCase) from the path.
      // LLMs may pass "message-id" (kebab) — we normalize so both forms work.
      const camelCaseParamName = paramName.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());

      // Look up param definition using normalized name (without $) for OData params,
      // or camelCase equivalent for kebab-case path params
      const paramDef = parameterDefinitions.find(
        (p) =>
          p.name === paramName ||
          p.name === camelCaseParamName ||
          (isOdataParam && p.name === normalizedParamName)
      );

      if (paramDef) {
        switch (paramDef.type) {
          case 'Path': {
            // Check if this parameter should skip URL encoding (for function-style API calls)
            const shouldSkipEncoding = config?.skipEncoding?.includes(paramName) ?? false;
            // Use encodeURIComponent but preserve '=' which is valid in path segments (RFC 3986)
            // and commonly appears in Microsoft Graph base64-encoded resource IDs.
            // Without this, IDs like "AAMk...AAA=" become "AAMk...AAA%3D" causing 404 errors.
            // First we encode, then unencode. Crazy, check out https://github.com/Softeria/ms-365-mcp-server/issues/245
            const encodedValue = shouldSkipEncoding
              ? (paramValue as string)
              : encodeURIComponent(paramValue as string).replace(/%3D/g, '=');

            // Replace both the original param name and the camelCase variant
            // to handle {message-id} (endpoints.json) and :messageId (generated client) formats
            path = path
              .replace(`{${paramName}}`, encodedValue)
              .replace(`:${paramName}`, encodedValue)
              .replace(`{${camelCaseParamName}}`, encodedValue)
              .replace(`:${camelCaseParamName}`, encodedValue);
            break;
          }

          case 'Query':
            if (paramValue !== '' && paramValue != null) {
              queryParams[fixedParamName] = `${paramValue}`;
            }
            break;

          case 'Body':
            if (paramDef.schema) {
              const parseResult = paramDef.schema.safeParse(paramValue);
              if (!parseResult.success) {
                const wrapped = { [paramName]: paramValue };
                const wrappedResult = paramDef.schema.safeParse(wrapped);
                if (wrappedResult.success) {
                  logger.info(
                    `Auto-corrected parameter '${paramName}': AI passed nested field directly, wrapped it as {${paramName}: ...}`
                  );
                  body = wrapped;
                } else {
                  body = paramValue;
                }
              } else {
                body = paramValue;
              }
            } else {
              body = paramValue;
            }
            break;

          case 'Header':
            headers[fixedParamName] = `${paramValue}`;
            break;
        }
      } else if (paramName === 'body') {
        body = paramValue;
        logger.info(`Set body param: ${JSON.stringify(body)}`);
      } else if (
        path.includes(`:${paramName}`) ||
        path.includes(`{${paramName}}`) ||
        path.includes(`:${camelCaseParamName}`) ||
        path.includes(`{${camelCaseParamName}}`)
      ) {
        // Fallback: path param not declared in tool.parameters (generated client omits them).
        // Replace placeholder directly so the URL is valid.
        const encodedValue = encodeURIComponent(paramValue as string).replace(/%3D/g, '=');
        path = path
          .replace(`{${paramName}}`, encodedValue)
          .replace(`:${paramName}`, encodedValue)
          .replace(`{${camelCaseParamName}}`, encodedValue)
          .replace(`:${camelCaseParamName}`, encodedValue);
        logger.info(`Path param fallback: replaced :${camelCaseParamName} with encoded value`);
      } else if (isOdataParam) {
        // Fallback: OData param recognised by name but absent from generated client's parameter
        // list — forward it as a query param rather than silently dropping it.
        queryParams[fixedParamName] = `${paramValue}`;
        logger.info(`OData param fallback: forwarded ${fixedParamName}=${paramValue}`);
      } else if (
        bodyShape &&
        (hasOwn(bodyShape, paramName) || hasOwn(bodyShape, camelCaseParamName)) &&
        !READ_ONLY_BODY_FIELDS.has(hasOwn(bodyShape, paramName) ? paramName : camelCaseParamName)
      ) {
        // Client flattened the body object into top-level params - rescue instead of
        // dropping. The read-only check uses the resolved name so kebab-case variants
        // can't sneak past
        const fieldName = hasOwn(bodyShape, paramName) ? paramName : camelCaseParamName;
        strayBodyFields[fieldName] = paramValue;
        logger.info(
          `Body field fallback: merging top-level param '${fieldName}' into request body`
        );
      } else {
        logger.warn(`Dropping unrecognized parameter '${paramName}' for tool ${tool.alias}`);
      }
    }

    if (Object.keys(strayBodyFields).length > 0) {
      if (isPlainObject(body)) {
        // If none of body's keys are schema fields but the schema has a `body` field
        // (message.body), the client meant it as that field - nest it. Spread order lets
        // an explicit body win over stray duplicates in both branches
        const keys = Object.keys(body);
        const bodyIsNestedField =
          bodyShape != null &&
          hasOwn(bodyShape, 'body') &&
          keys.length > 0 &&
          keys.every((k) => !hasOwn(bodyShape, k));
        body = bodyIsNestedField ? { ...strayBodyFields, body } : { ...strayBodyFields, ...body };
        logger.info(`Merged flattened body fields: ${Object.keys(strayBodyFields).join(', ')}`);
      } else if (body == null) {
        body = strayBodyFields;
        logger.info(`Merged flattened body fields: ${Object.keys(strayBodyFields).join(', ')}`);
      } else {
        logger.warn(
          `Cannot merge flattened body fields (${Object.keys(strayBodyFields).join(', ')}) into non-object request body; dropping them`
        );
      }
    }

    // Defense-in-depth: the calendar delta tools don't support $top (see
    // TOP_UNSUPPORTED_DELTA_TOOLS). Their user-facing schema strips top/$top, so
    // freshly-connected clients can't send it. Cached/stale clients (and ad-hoc
    // callers) might still try — drop it server-side before clamping or sending.
    if (TOP_UNSUPPORTED_DELTA_TOOLS.has(tool.alias)) {
      delete queryParams['$top'];
    }

    clampTopQueryParam(queryParams);
    normalizeSearchQueryParam(queryParams, tool.path);

    const preferValues: string[] = [];

    // Handle timezone parameter for calendar endpoints
    if (config?.supportsTimezone && params.timezone) {
      preferValues.push(`outlook.timezone="${params.timezone}"`);
      logger.info(`Setting timezone preference: outlook.timezone="${params.timezone}"`);
    }

    const bodyFormat = process.env.MS365_MCP_BODY_FORMAT || 'text';
    if (bodyFormat !== 'html' && tool.method.toUpperCase() === 'GET') {
      preferValues.push(`outlook.body-content-type="${bodyFormat}"`);
    }

    if (preferValues.length > 0) {
      headers['Prefer'] = preferValues.join(', ');
    }

    // Handle expandExtendedProperties parameter for calendar endpoints
    if (config?.supportsExpandExtendedProperties && params.expandExtendedProperties === true) {
      const expandValue = 'singleValueExtendedProperties';
      if (queryParams['$expand']) {
        queryParams['$expand'] += `,${expandValue}`;
      } else {
        queryParams['$expand'] = expandValue;
      }
      logger.info(`Adding $expand=${expandValue} for extended properties`);
    }

    if (config?.contentType) {
      headers['Content-Type'] = config.contentType;
      logger.info(`Setting custom Content-Type: ${config.contentType}`);
    }

    if (config?.acceptType) {
      headers['Accept'] = config.acceptType;
      logger.info(`Setting custom Accept: ${config.acceptType}`);
    }

    if (Object.keys(queryParams).length > 0) {
      const queryString = Object.entries(queryParams)
        .map(([key, value]) => `${key}=${encodeURIComponent(value).replace(/%2C/gi, ',')}`)
        .join('&');
      path = `${path}${path.includes('?') ? '&' : '?'}${queryString}`;
    }

    const options: {
      method: string;
      headers: Record<string, string>;
      body?: string | Buffer | Uint8Array;
      rawResponse?: boolean;
      includeHeaders?: boolean;
      excludeResponse?: boolean;
      queryParams?: Record<string, string>;
      accessToken?: string;
      apiVersion?: string;
      forceJsonOutput?: boolean;
    } = {
      method: tool.method.toUpperCase(),
      headers,
    };

    // Route beta-flagged endpoints to the /beta surface; everything else stays on v1.0.
    if (config?.apiVersion) {
      options.apiVersion = config.apiVersion;
    }

    if (options.method !== 'GET' && body) {
      if (tool.requestFormat === 'binary' && typeof body === 'string') {
        options.body = Buffer.from(body, 'base64');
        if (!config?.contentType) {
          headers['Content-Type'] = 'application/octet-stream';
        }
      } else if (config?.contentType === 'text/html') {
        if (typeof body === 'string') {
          options.body = body;
        } else if (typeof body === 'object' && 'content' in body) {
          options.body = (body as { content: string }).content;
        } else {
          options.body = String(body);
        }
      } else {
        options.body = typeof body === 'string' ? body : JSON.stringify(body);
      }
    }

    const isProbablyMediaContent =
      tool.errors?.some((error) => error.description === 'Retrieved media content') ||
      path.endsWith('/content');

    if (config?.returnDownloadUrl && path.endsWith('/content')) {
      path = path.replace(/\/content$/, '');
      logger.info(
        `Auto-returning download URL for ${tool.alias} (returnDownloadUrl=true in endpoints.json)`
      );
    } else if (isProbablyMediaContent) {
      options.rawResponse = true;
    }

    // Set includeHeaders if requested
    if (params.includeHeaders === true) {
      options.includeHeaders = true;
    }

    // Set excludeResponse if requested
    if (params.excludeResponse === true) {
      options.excludeResponse = true;
    }

    // Pass account-resolved token if available
    if (accountAccessToken) {
      options.accessToken = accountAccessToken;
    }

    targetResource = deriveTargetResource({
      pathPattern: config?.pathPattern ?? tool.path,
      params,
    });

    // Redact accessToken from log output to prevent credential leakage
    const { accessToken: _redacted, ...safeOptions } = options;
    logger.info(
      `Making graph request to ${path} with options: ${JSON.stringify(safeOptions)}${_redacted ? ' [accessToken=REDACTED]' : ''}`
    );

    const fetchAllPages = params.fetchAllPages === true;
    const paginationEnabled = paginationAllowed();
    if (fetchAllPages && !paginationEnabled) {
      logger.info(
        'fetchAllPages requested but MS365_MCP_ALLOW_PAGINATION is disabled; returning first page only'
      );
    }
    // Force every page to JSON so the merge loop can parse them. Under --toon they'd
    // be TOON and JSON.parse would throw, silently returning only page one (#560).
    // The merged result gets re-encoded once at the end.
    const mergePages = fetchAllPages && paginationEnabled;
    if (mergePages) {
      options.forceJsonOutput = true;
    }

    let response = await graphClient.graphRequest(path, options);

    if (mergePages && response?.content?.[0]?.text) {
      type ODataPage = {
        value?: unknown[];
        '@odata.nextLink'?: string;
        '@odata.deltaLink'?: string;
        '@odata.count'?: number;
        [key: string]: unknown;
      };
      let combinedResponse: ODataPage | undefined;
      try {
        combinedResponse = JSON.parse(response.content[0].text) as ODataPage;

        // Only merge if page one is actually a collection. fetchAllPages can be set
        // on a single-object GET too, and we'd otherwise graft a bogus value:[] on it.
        const firstValue = combinedResponse.value;
        if (Array.isArray(firstValue)) {
          let allItems: unknown[] = firstValue;
          let nextLink = combinedResponse['@odata.nextLink'];
          let pageCount = 1;
          const maxPages = positiveIntFromEnv('MS365_MCP_MAX_PAGES', DEFAULT_MAX_PAGES);
          const maxItems = positiveIntFromEnv('MS365_MCP_MAX_ITEMS', DEFAULT_MAX_ITEMS);
          // Graph only emits @odata.deltaLink on the final page of a /delta query.
          // Track it across the pagination loop so we can stamp it on the combined
          // response — otherwise fetchAllPages on a /delta endpoint silently drops
          // the resume token and forces callers to re-list from scratch.
          let deltaLink = combinedResponse['@odata.deltaLink'];

          while (nextLink && pageCount < maxPages && allItems.length < maxItems) {
            logger.info(`Fetching page ${pageCount + 1} from: ${nextLink}`);

            // Extract path + query string from the nextLink URL.
            // Pass the full path (with query string) as the endpoint so that
            // $skiptoken and other pagination params are preserved.
            // Previously, query params were extracted into nextOptions.queryParams
            // but graphRequest/performRequest never read that field — they were lost.
            const url = new URL(nextLink);
            // nextLink is absolute and version-qualified (/v1.0/... or /beta/...). Strip the
            // version segment so performRequest can re-apply the request's own apiVersion.
            const nextPath = url.pathname.replace(/^\/(v1\.0|beta)/, '') + url.search;
            const nextOptions = { ...options };

            const nextResponse = await graphClient.graphRequest(nextPath, nextOptions);
            if (nextResponse?.content?.[0]?.text) {
              const nextJsonResponse = JSON.parse(nextResponse.content[0].text) as ODataPage;
              if (Array.isArray(nextJsonResponse.value)) {
                allItems = allItems.concat(nextJsonResponse.value);
              }
              nextLink = nextJsonResponse['@odata.nextLink'];
              if (nextJsonResponse['@odata.deltaLink']) {
                deltaLink = nextJsonResponse['@odata.deltaLink'];
              }
              pageCount++;
            } else {
              break;
            }
          }

          if (pageCount >= maxPages) {
            logger.warn(`Reached maximum page limit (${maxPages}) for pagination`);
          }
          if (allItems.length >= maxItems) {
            logger.warn(
              `Reached maximum item limit (${maxItems}) for pagination — truncated at ${allItems.length} items`
            );
          }

          combinedResponse.value = allItems;
          if (combinedResponse['@odata.count']) {
            combinedResponse['@odata.count'] = allItems.length;
          }
          delete combinedResponse['@odata.nextLink'];
          if (deltaLink) {
            combinedResponse['@odata.deltaLink'] = deltaLink;
          }

          logger.info(
            `Pagination complete: collected ${allItems.length} items across ${pageCount} pages`
          );
        }
      } catch (e) {
        logger.error(`Error during pagination: ${e}`);
      }

      // Re-encode once in the configured format. Runs whenever page one parsed
      // (non-collection skip and mid-loop abort included), so a --toon client
      // never gets handed the forced-JSON body.
      if (combinedResponse !== undefined) {
        response.content[0].text = graphClient.serialize(combinedResponse);
      }
    }

    if (response?.content?.[0]?.text) {
      const responseText = response.content[0].text;
      logger.info(`Response size: ${responseText.length} characters`);

      try {
        const jsonResponse = JSON.parse(responseText);
        if (jsonResponse.value && Array.isArray(jsonResponse.value)) {
          logger.info(`Response contains ${jsonResponse.value.length} items`);
        }
        if (jsonResponse['@odata.nextLink']) {
          logger.info(`Response has pagination nextLink: ${jsonResponse['@odata.nextLink']}`);
        }
      } catch {
        // Non-JSON response
      }
    }

    // Convert McpResponse to CallToolResult with the correct structure
    const content: ContentItem[] = response.content.map((item) => ({
      type: 'text' as const,
      text: item.text,
    }));

    auditLog({
      event: 'tool.call',
      request_id: requestId,
      user_principal_name: upn,
      tool: tool.alias,
      http_method: httpMethod,
      status: response.isError ? 'error' : 'success',
      duration_ms: Date.now() - startTime,
      ...(targetResource ? { target_resource: targetResource } : {}),
    });

    return {
      content,
      _meta: response._meta,
      isError: response.isError,
    };
  } catch (error) {
    const err = error as { name?: string; code?: string | number; status?: string | number };
    logger.error(`Error in tool ${tool.alias}: ${(error as Error).message}`);
    auditLog({
      event: 'tool.call',
      request_id: requestId,
      user_principal_name: upn,
      tool: tool.alias,
      http_method: httpMethod,
      status: 'error',
      duration_ms: Date.now() - startTime,
      ...(targetResource ? { target_resource: targetResource } : {}),
      error_type: err?.name || 'Error',
      error_code: err?.status ?? err?.code,
    });
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            error: `Error in tool ${tool.alias}: ${(error as Error).message}`,
          }),
        },
      ],
      isError: true,
    };
  }
}

export function registerGraphTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  enabledToolsPattern?: string,
  orgMode: boolean = false,
  authManager?: AuthManager,
  multiAccount: boolean = false,
  accountNames: string[] = [],
  allowedScopesValue?: string,
  httpMode: boolean = false,
  attachmentProxy: boolean = false
): number {
  let enabledToolsRegex: RegExp | undefined;
  if (enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, 'i');
      logger.info(`Tool filtering enabled with pattern: ${enabledToolsPattern}`);
    } catch {
      logger.error(`Invalid tool filter regex pattern: ${enabledToolsPattern}. Ignoring filter.`);
    }
  }

  let registeredCount = 0;
  let skippedCount = 0;
  let failedCount = 0;
  const allowedScopes = parseAllowedScopes(allowedScopesValue);
  const disabledByAllowedScopes: DisabledToolScope[] = [];

  for (const tool of allEndpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);
    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      logger.info(`Skipping work account tool ${tool.alias} - not in org mode`);
      skippedCount++;
      continue;
    }

    if (attachmentProxy && isProxySuppressedGraphTool(tool.method, endpointConfig?.pathPattern)) {
      logger.info(`Skipping raw-byte tool ${tool.alias} - --attachment-proxy is set`);
      skippedCount++;
      continue;
    }

    const method = tool.method.toUpperCase();
    if (readOnly && method !== 'GET') {
      // Allow POST endpoints that are explicitly marked as readOnly in endpoints.json
      // (e.g. get-schedule, find-meeting-times which are read-only queries via POST).
      // PATCH/DELETE are always blocked in read-only mode.
      if (!(method === 'POST' && endpointConfig?.readOnly)) {
        logger.info(`Skipping write operation ${tool.alias} in read-only mode`);
        skippedCount++;
        continue;
      }
    }

    if (enabledToolsRegex && !enabledToolsRegex.test(tool.alias)) {
      logger.info(`Skipping tool ${tool.alias} - doesn't match filter pattern`);
      skippedCount++;
      continue;
    }

    const missingScopes =
      allowedScopes !== undefined && !endpointConfig
        ? ['endpoint scope metadata']
        : getMissingAllowedScopesForGroups(
            getEndpointScopeGroups(endpointConfig, orgMode),
            allowedScopes
          );
    if (missingScopes.length > 0) {
      disabledByAllowedScopes.push({ toolName: tool.alias, missingScopes });
      skippedCount++;
      continue;
    }

    const paramSchema: Record<string, z.ZodTypeAny> = {};
    if (tool.parameters && tool.parameters.length > 0) {
      for (const param of tool.parameters) {
        // Lenient Body validation, or the SDK strips a flattened body value to {} (#569)
        paramSchema[param.name] =
          param.type === 'Body' && param.schema
            ? lenientBodySchema(param.schema as z.ZodTypeAny)
            : param.schema || z.any();
      }
    }

    // Extract path parameters from the path pattern (e.g., :todoTaskListId from /me/todo/lists/:todoTaskListId/tasks)
    // The generated client omits these from tool.parameters, so we add them manually.
    const pathParamMatches = tool.path.matchAll(/:([a-zA-Z]+)/g);
    for (const match of pathParamMatches) {
      const pathParamName = match[1];
      if (!(pathParamName in paramSchema)) {
        paramSchema[pathParamName] = z.string().describe(describePathParam(pathParamName));
      }
    }

    if (isFetchAllPagesApplicable(tool)) {
      const maxPages = getMaxPages();
      paramSchema['fetchAllPages'] = z
        .boolean()
        .describe(getFetchAllPagesParamDescription(maxPages))
        .optional();
    }

    // Override OData parameter descriptions with spec-gap guidance. Text lives in
    // lib/param-descriptions.ts, shared with describeToolSchema (--discovery mode),
    // so the two paths cannot describe the same parameter differently.
    if (paramSchema['filter'] !== undefined || paramSchema['$filter'] !== undefined) {
      const key = paramSchema['$filter'] !== undefined ? '$filter' : 'filter';
      paramSchema[key] = z.string().describe(FILTER_PARAM_DESCRIPTION).optional();
    }
    if (paramSchema['search'] !== undefined || paramSchema['$search'] !== undefined) {
      const key = paramSchema['$search'] !== undefined ? '$search' : 'search';
      paramSchema[key] = z.string().describe(SEARCH_PARAM_DESCRIPTION).optional();
    }
    if (paramSchema['select'] !== undefined || paramSchema['$select'] !== undefined) {
      const key = paramSchema['$select'] !== undefined ? '$select' : 'select';
      paramSchema[key] = z.string().describe(SELECT_PARAM_DESCRIPTION).optional();
    }
    // The spec describes every $expand as "Expand related entities", which says nothing about
    // what is expandable. Models pass non-navigation properties — message body is the one I
    // hit repeatedly — and Graph answers 400 "Parsing OData Select and Expand failed".
    // Restated as the override rather than a new schema: $expand is already array<string>
    // everywhere, so the type is unchanged in practice.
    if (paramSchema['expand'] !== undefined || paramSchema['$expand'] !== undefined) {
      const key = paramSchema['$expand'] !== undefined ? '$expand' : 'expand';
      paramSchema[key] = z.array(z.string()).describe(EXPAND_PARAM_DESCRIPTION).optional();
    }
    if (paramSchema['orderby'] !== undefined || paramSchema['$orderby'] !== undefined) {
      const key = paramSchema['$orderby'] !== undefined ? '$orderby' : 'orderby';
      paramSchema[key] = z.string().describe(ORDERBY_PARAM_DESCRIPTION).optional();
    }
    // The calendar delta tools don't support $top (see TOP_UNSUPPORTED_DELTA_TOOLS) —
    // page size is controlled via Prefer: odata.maxpagesize. Strip top/$top from
    // their schemas so callers can't reach for a parameter that won't work. Other
    // delta tools (message/driveItem/site) do support $top, so leave them alone.
    // Server-side defense-in-depth in executeGraphTool handles stale clients.
    if (shouldOmitTopParam(tool.alias)) {
      delete paramSchema['top'];
      delete paramSchema['$top'];
    } else if (paramSchema['top'] !== undefined || paramSchema['$top'] !== undefined) {
      const key = paramSchema['$top'] !== undefined ? '$top' : 'top';
      paramSchema[key] = z.number().describe(TOP_PARAM_DESCRIPTION).optional();
    }
    if (paramSchema['skip'] !== undefined || paramSchema['$skip'] !== undefined) {
      const key = paramSchema['$skip'] !== undefined ? '$skip' : 'skip';
      paramSchema[key] = z.number().describe(SKIP_PARAM_DESCRIPTION).optional();
    }
    if (paramSchema['count'] !== undefined || paramSchema['$count'] !== undefined) {
      const countKey = paramSchema['$count'] !== undefined ? '$count' : 'count';
      paramSchema[countKey] = z.boolean().describe(COUNT_PARAM_DESCRIPTION).optional();
    }

    // Add account parameter for multi-account mode.
    // Layer 2: Account names are surfaced in the description (not as a strict enum) so the LLM
    // sees available accounts upfront without a round-trip, but accounts added mid-session via
    // --login are still accepted — getTokenForAccount() handles validation at runtime.
    if (multiAccount) {
      paramSchema['account'] = z
        .string()
        .describe(getAccountParamDescription(accountNames))
        .optional();
    }

    // Add includeHeaders parameter for all tools to capture ETags and other headers
    paramSchema['includeHeaders'] = z
      .boolean()
      .describe('Include response headers (including ETag) in the response metadata')
      .optional();

    // Add excludeResponse parameter to only return success/failure indication
    paramSchema['excludeResponse'] = z
      .boolean()
      .describe('Exclude the full response body and only return success or failure indication')
      .optional();

    // Destructive tools (POST except readOnly, PATCH, PUT, DELETE) require an
    // explicit `confirm: true` server-side gate. See isDestructiveOperation +
    // executeGraphTool for the enforcement; surface the param in the schema so
    // the LLM/agent sees it upfront.
    const destructive = isDestructiveOperation(tool.method, endpointConfig);
    if (destructive) {
      paramSchema['confirm'] = z.boolean().describe(CONFIRM_PARAM_DESCRIPTION).optional();
    }

    // Add timezone parameter for calendar endpoints that support it
    if (endpointConfig?.supportsTimezone) {
      paramSchema['timezone'] = z.string().describe(TIMEZONE_PARAM_DESCRIPTION).optional();
    }

    // Add expandExtendedProperties parameter for calendar endpoints that support it
    if (endpointConfig?.supportsExpandExtendedProperties) {
      paramSchema['expandExtendedProperties'] = z
        .boolean()
        .describe(EXPAND_EXTENDED_PROPERTIES_PARAM_DESCRIPTION)
        .optional();
    }

    // Build the tool description, optionally appending LLM tips
    let toolDescription = withApiVersionPrefix(
      (endpointConfig?.descriptionOverride ?? tool.description) ||
        `Execute ${tool.method.toUpperCase()} request to ${tool.path}`,
      endpointConfig
    );
    if (endpointConfig?.llmTip) {
      toolDescription += `\n\n💡 TIP: ${endpointConfig.llmTip}`;
    }

    // An endpoint marked readOnly in endpoints.json (e.g. a POST query like
    // copilot-retrieve) is a read-only operation despite its write verb, so derive
    // the hints from that flag rather than the HTTP method alone — otherwise a
    // read-only query lands as destructiveHint:true and clients mis-rank it.
    const isReadOnlyTool = tool.method.toUpperCase() === 'GET' || endpointConfig?.readOnly === true;

    try {
      // .passthrough() object, not a raw shape - the SDK wraps raw shapes in z.object()
      // and strips unknown keys before the handler runs, which is exactly how #569's
      // flattened subject/toRecipients got lost
      server.registerTool(
        tool.alias,
        {
          title: tool.alias,
          description: toolDescription,
          inputSchema: z.object(paramSchema).passthrough(),
          annotations: {
            title: tool.alias,
            readOnlyHint: isReadOnlyTool,
            destructiveHint: destructive,
            openWorldHint: true, // All tools call Microsoft Graph API
          },
        },
        async (params: Record<string, unknown>) =>
          executeGraphTool(tool, endpointConfig, graphClient, params, authManager)
      );
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool ${tool.alias}: ${(error as Error).message}`);
      failedCount++;
    }
  }

  if (multiAccount) {
    logger.info('Multi-account mode: "account" parameter injected into all tool schemas');
  }

  if (disabledByAllowedScopes.length > 0) {
    logger.info(
      `Allowed scopes disabled ${disabledByAllowedScopes.length} Graph tools: ${formatDisabledToolsForLog(disabledByAllowedScopes)}`
    );
  }

  const utilityCtx: UtilityToolContext = {
    graphClient,
    authManager,
    multiAccount,
    accountNames,
  };
  for (const utility of selectUtilityTools({
    readOnly,
    httpMode,
    enabledTools: enabledToolsPattern,
    attachmentProxy,
  })) {
    try {
      registerUtilityToolWithMcp(server, utility, utilityCtx);
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool ${utility.name}: ${(error as Error).message}`);
      failedCount++;
    }
  }

  // Layer 3 (list-accounts tool) is registered by registerAuthTools in auth-tools.ts.
  // It is the canonical owner of account discovery — no duplicate registration here.

  logger.info(
    `Tool registration complete: ${registeredCount} registered, ${skippedCount} skipped, ${failedCount} failed`
  );
  return registeredCount;
}

export function buildToolsRegistry(
  readOnly: boolean,
  orgMode: boolean,
  enabledToolsRegex?: RegExp,
  allowedScopesValue?: string,
  disabledByAllowedScopes: Array<{ toolName: string; missingScopes: string[] }> = [],
  attachmentProxy: boolean = false
): Map<string, { tool: (typeof api.endpoints)[0]; config: EndpointConfig | undefined }> {
  const toolsMap = new Map<
    string,
    { tool: (typeof api.endpoints)[0]; config: EndpointConfig | undefined }
  >();
  const allowedScopes = parseAllowedScopes(allowedScopesValue);

  for (const tool of allEndpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);

    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      continue;
    }

    if (attachmentProxy && isProxySuppressedGraphTool(tool.method, endpointConfig?.pathPattern)) {
      continue;
    }

    const method = tool.method.toUpperCase();
    if (readOnly && method !== 'GET') {
      if (!(method === 'POST' && endpointConfig?.readOnly)) {
        continue;
      }
    }

    if (enabledToolsRegex && !enabledToolsRegex.test(tool.alias)) {
      continue;
    }

    const missingScopes =
      allowedScopes !== undefined && !endpointConfig
        ? ['endpoint scope metadata']
        : getMissingAllowedScopesForGroups(
            getEndpointScopeGroups(endpointConfig, orgMode),
            allowedScopes
          );
    if (missingScopes.length > 0) {
      disabledByAllowedScopes.push({ toolName: tool.alias, missingScopes });
      continue;
    }

    toolsMap.set(tool.alias, { tool, config: endpointConfig });
  }

  return toolsMap;
}

/**
 * Builds a BM25 index over the tool registry. Name tokens are weighted 3x and llmTip
 * tokens 2x via repetition, so a tool whose name matches the query outranks one that
 * merely mentions the query term in its Microsoft-supplied description.
 */
export function buildDiscoverySearchIndex(
  toolsRegistry: ReturnType<typeof buildToolsRegistry>,
  utilityTools: readonly UtilityTool[] = []
): DiscoverySearchIndex {
  // Cap contribution from the `description` and `llmTip` fields so a verbose llmTip
  // (e.g. the KQL search-syntax guide on list-mail-messages, ~300 tokens) doesn't
  // inflate a tool's doc length and crush BM25's length normalization. Names and
  // paths are short and reliable, so they stay uncapped and are repeated to carry
  // the bulk of the ranking signal. Tip excerpt (12 tokens) is enough to capture
  // the first "what this tool does" phrase without swamping the doc.
  const TIP_EXCERPT_TOKENS = 12;
  const DESC_CAP_TOKENS = 40;
  const docs: Array<{ id: string; tokens: string[] }> = [];
  const nameTokens = new Map<string, Set<string>>();
  for (const [name, { tool, config }] of toolsRegistry) {
    const nt = tokenize(name);
    nameTokens.set(name, new Set(nt));
    const pathTokens = tokenize(tool.path);
    const descTokens = tokenize(config?.descriptionOverride ?? tool.description).slice(
      0,
      DESC_CAP_TOKENS
    );
    const tipTokens = tokenize(config?.llmTip).slice(0, TIP_EXCERPT_TOKENS);
    const tokens = [
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...pathTokens,
      ...pathTokens,
      ...tipTokens,
      ...descTokens,
    ];
    docs.push({ id: name, tokens });
  }
  for (const utility of utilityTools) {
    const nt = tokenize(utility.name);
    nameTokens.set(utility.name, new Set(nt));
    const pathTokens = tokenize(utility.path);
    const keywordTokens = tokenize(utility.searchKeywords);
    const descTokens = tokenize(utility.description).slice(0, DESC_CAP_TOKENS);
    const tokens = [
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...pathTokens,
      ...pathTokens,
      ...keywordTokens,
      ...keywordTokens,
      ...descTokens,
    ];
    docs.push({ id: utility.name, tokens });
  }
  return { bm25: buildBM25Index(docs), nameTokens };
}

/**
 * BM25 + a "name precision" bonus: reward tools whose names contain a high fraction
 * of the query tokens (and consist mostly of query-matching tokens). This counteracts
 * cases where a tool with a longer or more off-topic description outranks a tool
 * whose name directly matches — a common problem because many endpoint descriptions
 * are the wrong Graph prose pasted in.
 */
export function scoreDiscoveryQuery(
  query: string,
  index: DiscoverySearchIndex
): Array<{ id: string; score: number }> {
  const queryTokenSet = new Set(tokenize(query));
  if (queryTokenSet.size === 0) return [];
  const ranked = scoreQuery(query, index.bm25);
  const NAME_BONUS_WEIGHT = 2;
  for (const r of ranked) {
    const nt = index.nameTokens.get(r.id);
    if (!nt || nt.size === 0) continue;
    let matchedIdf = 0;
    let matchedCount = 0;
    for (const qt of queryTokenSet) {
      if (nt.has(qt)) {
        matchedCount++;
        matchedIdf += index.bm25.idf.get(qt) ?? 0;
      }
    }
    if (matchedCount === 0) continue;
    const precision = matchedCount / nt.size;
    r.score += precision * matchedIdf * NAME_BONUS_WEIGHT;
  }
  ranked.sort((a, b) => b.score - a.score);
  return ranked;
}

export function registerDiscoveryTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  orgMode: boolean = false,
  authManager?: AuthManager,
  multiAccount: boolean = false,
  accountNames: string[] = [],
  enabledTools?: string,
  allowedScopesValue?: string,
  httpMode: boolean = false,
  attachmentUrls: boolean = false,
  attachmentProxy: boolean = false
): void {
  let enabledToolsRegex: RegExp | undefined;
  if (enabledTools) {
    try {
      enabledToolsRegex = new RegExp(enabledTools, 'i');
      logger.info(`Discovery mode: filtering tools with pattern ${enabledTools}`);
    } catch (error) {
      logger.error(
        `Invalid --enabled-tools regex ${JSON.stringify(enabledTools)} — ignoring filter: ${(error as Error).message}`
      );
    }
  }

  const disabledByAllowedScopes: Array<{ toolName: string; missingScopes: string[] }> = [];
  const toolsRegistry = buildToolsRegistry(
    readOnly,
    orgMode,
    enabledToolsRegex,
    allowedScopesValue,
    disabledByAllowedScopes,
    attachmentProxy
  );
  if (disabledByAllowedScopes.length > 0) {
    logger.info(
      `Discovery mode: allowed scopes disabled ${disabledByAllowedScopes.length} Graph tools: ${formatDisabledToolsForLog(disabledByAllowedScopes)}`
    );
  }
  const utilityTools = selectUtilityTools({ readOnly, httpMode, enabledTools, attachmentProxy });
  const searchIndex = buildDiscoverySearchIndex(toolsRegistry, utilityTools);
  const totalCount = toolsRegistry.size + utilityTools.length;
  logger.info(
    `Discovery mode: ${totalCount} tools (${toolsRegistry.size} Graph + ${utilityTools.length} utility)`
  );

  const utilityCtx: UtilityToolContext = {
    graphClient,
    authManager,
    multiAccount,
    accountNames,
  };
  const utilityByName = new Map(utilityTools.map((u) => [u.name, u]));

  const categoryNames = Object.keys(TOOL_CATEGORIES).join(', ');

  const toResultEntry = (name: string) => {
    const entry = toolsRegistry.get(name);
    if (entry) {
      const { tool, config } = entry;
      return {
        name,
        method: tool.method.toUpperCase(),
        path: tool.path,
        description: withApiVersionPrefix(
          (config?.descriptionOverride ?? tool.description) ||
            `${tool.method.toUpperCase()} ${tool.path}`,
          config
        ),
        ...(config?.llmTip ? { llmTip: config.llmTip } : {}),
      };
    }
    const utility = utilityByName.get(name);
    if (utility) {
      return {
        name: utility.name,
        method: utility.method,
        path: utility.path,
        description: utility.description,
      };
    }
    return null;
  };

  server.tool(
    'search-tools',
    `Search through ${totalCount} tools (${toolsRegistry.size} Microsoft Graph API operations + ${utilityTools.length} server utilities like download-bytes). Ranks results by BM25 over tool name, llmTip, description, and path. After picking a tool, call get-tool-schema for parameters, then execute-tool.`,
    {
      query: z
        .string()
        .describe(
          'Natural-language query. Tokenized and BM25-ranked. E.g. "send email", "download photo", "list unread messages".'
        )
        .optional(),
      category: z.string().describe(`Optional pre-filter by category: ${categoryNames}`).optional(),
      limit: z.number().describe('Maximum results (default: 10, max: 50)').optional(),
    },
    {
      title: 'search-tools',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ query, category, limit = 10 }) => {
      const maxLimit = Math.min(Math.max(limit, 1), 50);
      // Built with the same flags the registration filter used. A category pattern that ignores
      // them would hide a tool this server did register under that category — the same staleness
      // that made get-download-url unreachable, one layer up.
      const categoryPattern = category
        ? getCategoryPattern(category, { attachmentUrls, attachmentProxy })
        : undefined;
      const categoryFilter = (name: string) => !categoryPattern || categoryPattern.test(name);

      let orderedNames: string[];
      if (query && query.trim().length > 0) {
        const ranked = scoreDiscoveryQuery(query, searchIndex);
        orderedNames = ranked.map((r) => r.id).filter(categoryFilter);
      } else {
        orderedNames = [...toolsRegistry.keys(), ...utilityTools.map((u) => u.name)].filter(
          categoryFilter
        );
      }

      const tools = orderedNames.slice(0, maxLimit).map(toResultEntry).filter(Boolean);

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                found: tools.length,
                total: totalCount,
                tools,
                tip: 'Call get-tool-schema(tool_name) to see parameters before invoking execute-tool.',
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  server.tool(
    'get-tool-schema',
    'Returns the full parameter schema (name, placement, required, JSON Schema) for a tool discovered via search-tools. Call this before execute-tool so you know what parameters to pass and what enum values are valid.',
    {
      tool_name: z.string().describe('Exact tool name from search-tools (e.g. "send-mail")'),
    },
    {
      title: 'get-tool-schema',
      readOnlyHint: true,
      openWorldHint: false,
    },
    async ({ tool_name }) => {
      const entry = toolsRegistry.get(tool_name);
      if (entry) {
        const schema = describeToolSchema(entry.tool, entry.config, { multiAccount, accountNames });
        return {
          content: [{ type: 'text', text: JSON.stringify(schema, null, 2) }],
        };
      }
      const utility = utilityByName.get(tool_name);
      if (utility) {
        const schema = describeUtilityToolSchema(utility, utilityCtx);
        return {
          content: [{ type: 'text', text: JSON.stringify(schema, null, 2) }],
        };
      }
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              error: `Tool not found: ${tool_name}`,
              tip: 'Use search-tools to find available tools.',
            }),
          },
        ],
        isError: true,
      };
    }
  );

  server.tool(
    'execute-tool',
    'Execute a Microsoft Graph API tool by name. Workflow: search-tools → get-tool-schema → execute-tool. Call get-tool-schema first for any tool you have not seen before — passing the wrong shape to parameters will fail validation or return a Graph 400. For list endpoints, prefer modest $top plus $select.',
    {
      tool_name: z.string().describe('Name of the tool to execute (e.g., "list-mail-messages")'),
      parameters: z
        .record(z.any())
        .describe(
          'Parameters shaped per get-tool-schema. Path/query/header params go at the top level; request bodies go under "body".'
        )
        .optional(),
    },
    {
      title: 'execute-tool',
      readOnlyHint: false,
      destructiveHint: true,
      openWorldHint: true,
    },
    async ({ tool_name, parameters = {} }) => {
      const toolData = toolsRegistry.get(tool_name);
      if (toolData) {
        return executeGraphTool(
          toolData.tool,
          toolData.config,
          graphClient,
          parameters,
          authManager
        );
      }
      const utility = utilityByName.get(tool_name);
      if (utility) {
        return utility.execute(parameters, utilityCtx);
      }
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              error: `Tool not found: ${tool_name}`,
              tip: 'Use search-tools to find available tools.',
            }),
          },
        ],
        isError: true,
      };
    }
  );

  // Layer 3 (list-accounts) is registered by registerAuthTools — no duplicate here.
}
