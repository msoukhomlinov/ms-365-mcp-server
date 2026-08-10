/** Shared context for MCP `initialize.instructions` (hosts that forward it to the model). */
export type McpInstructionsContext = {
  orgMode: boolean;
  readOnly: boolean;
  multiAccount: boolean;
  /**
   * The tool names this server actually registered
   * (`resolveRegisteredToolNames`), not the flags that shaped them.
   *
   * Flags are one step removed from the truth and that gap shipped a bug:
   * `--attachment-proxy` was taken as proof read-document exists, but an
   * --enabled-tools filter can drop it while startup only warns, leaving the
   * instructions promising a tool the server never registered. Every clause
   * below that names a tool declares it and is emitted only if it is here.
   */
  registeredTools: ReadonlySet<string>;
};

/**
 * One piece of guidance, and the tools it directs the model to.
 *
 * Declaring the names next to the text is what makes the invariant hold without
 * anyone having to classify prose: a clause is emitted only where every tool it
 * names is registered. There is no "directive vs explanatory" judgement to get
 * wrong, because a sentence about a tool is worth saying exactly where that
 * tool can be called. The stdio-only download-bytes-to-file guidance survives
 * in stdio and goes over HTTP for that reason alone, not as a carve-out.
 *
 * Clauses naming auth tools (list-accounts) or the discovery meta-tools are
 * left undeclared: they are registered by other modules that
 * `registeredTools` does not cover, and the option that emits those clauses is
 * the same one that registers them.
 */
type GuidanceClause = { text: string; tools?: readonly string[] };

function emit(clauses: readonly GuidanceClause[], registered: ReadonlySet<string>): string[] {
  return clauses
    .filter((clause) => (clause.tools ?? []).every((tool) => registered.has(tool)))
    .map((clause) => clause.text);
}

/**
 * Byte/document guidance, one self-contained sentence per tool.
 *
 * Self-contained is the whole design: each sentence names exactly one tool and
 * survives its neighbours' removal, so any combination of registered byte tools
 * produces true text with no per-combination branch and no cross-reference to a
 * tool that may be gone. When none of them registered — proxy mode with a
 * filter that also dropped read-document — the paragraph says so instead of
 * naming a tool that is not there.
 */
function buildByteContentInstructions(registered: ReadonlySet<string>): string {
  const body = emit(
    [
      {
        tools: ['read-document'],
        text:
          'read documents with read-document, which returns markdown and takes a relative Microsoft ' +
          'Graph byte path — a mail or event attachment ' +
          '(/me/messages/{message-id}/attachments/{attachment-id}/$value; the /$value suffix is ' +
          'required), a raw message (/me/messages/{message-id}/$value), a drive or SharePoint file ' +
          '(/drives/{drive-id}/items/{driveItem-id}/content), or another authenticated /$value ' +
          'endpoint. Absolute URLs are not accepted, and pages, offset and maxChars read a long ' +
          'document in parts. This server registers no tool that returns raw bytes, base64 content, ' +
          'or a download URL.',
      },
      {
        tools: ['get-download-url'],
        text:
          'for large drive/SharePoint file content, prefer get-download-url to resolve a ' +
          'pre-authenticated URL for out-of-band download. It covers authenticated byte endpoints ' +
          'too — mail attachments, meeting recordings, other /$value paths — but only when this ' +
          'server runs with --enable-attachment-urls (HTTP mode, server-held credentials), where it ' +
          'mints a short-lived URL of its own that accepts a few fetches (so a converter may probe ' +
          'and then convert the same URL, and a failed fetch is retried on it rather than ' +
          're-minted); without the flag, or when Graph identity comes from the request (OAuth, OBO, ' +
          'or bearer mode), it refuses them.',
      },
      {
        tools: ['download-bytes'],
        text:
          'use download-bytes for authenticated byte reads such as mail attachments, profile photos, ' +
          'Teams hosted content, and meeting recordings, and whenever a mint is refused. It returns ' +
          'base64 in the tool response. These tools take relative Microsoft Graph paths, not ' +
          'absolute URLs.',
      },
      {
        tools: ['download-bytes-to-file'],
        // Leads with prose, not the tool name: clauses after the first are
        // capitalised when joined, and capitalising a tool name breaks it.
        text:
          'in stdio mode, download-bytes-to-file writes those same authenticated bytes straight to a ' +
          'local absolute path instead of returning base64 — the out-of-band option for large mail ' +
          'attachments and meeting recordings where no URL can be minted.',
      },
    ],
    registered
  );

  if (body.length === 0) {
    // The narrow claim, deliberately. These four tools are the binary and
    // document readers; their absence says nothing about Graph tools that
    // return content inside their JSON — get-mail-message returns the message
    // body either way — so "no content can be read" would be false and would
    // have the model refuse work this server can do. Not enumerated either: a
    // list of content-returning Graph tools goes stale against endpoints.json
    // every time upstream adds one, and claiming a read that is not there is
    // the failure this guidance exists to prevent.
    return (
      'Files / binary content: this server registers no tool that returns binary or attachment ' +
      'bytes, mints a download URL, or converts a document to text, so file content, attachment ' +
      'bytes and raw message MIME cannot be read here. Text a Graph tool already returns in its ' +
      'JSON response body — a message body, an item description — is unaffected.'
    );
  }
  // Each clause is written to follow the section label, so the first reads
  // correctly after the colon and the rest have to be capitalised: which clause
  // lands first depends on what registered, and any of them can. A hyphenated
  // leading token is a tool name (download-bytes-to-file) and is left alone --
  // capitalising it would rename the tool.
  const [lead, ...rest] = body;
  const capitalize = (s: string) =>
    /^[a-z0-9]+-/.test(s) ? s : s.charAt(0).toUpperCase() + s.slice(1);
  return `Files / binary content: ${[lead, ...rest.map(capitalize)].join(' ')}`;
}

function buildGeneralMcpInstructions(opts: McpInstructionsContext): string {
  const parts = emit(
    [
      {
        text: 'Microsoft 365 MCP exposes Microsoft Graph through MCP tools. Use each tool name, description, and parameter schema as the source of truth.',
      },
      {
        text: 'Microsoft Graph OData: do not combine $filter with $search on the same request. For lists, prefer modest $top (or top) and $select; avoid very large pages unless the user needs them.',
      },
      {
        text: 'Mail and message $search uses KQL; the $search query parameter value must be double-quoted per Graph (see search-query-parameter in Microsoft Graph docs).',
      },
      {
        tools: ['list-users'],
        text: 'When you need an organizational user or recipient address, resolve it with list-users (or another directory tool); do not invent SMTP addresses.',
      },
      {
        text: 'Directory $search on collections such as /users or /groups requires ConsistencyLevel: eventual when the tool exposes that header.',
      },
      {
        text: 'Teams chat and channel messages: prefer HTML contentType in the body; plain text is often mangled by Graph.',
      },
      { text: buildByteContentInstructions(opts.registeredTools) },
      {
        tools: ['upload-file-content'],
        text: 'For uploads, upload-file-content takes a base64 string body up to 4MB.',
      },
      {
        tools: ['create-upload-session'],
        text: 'For uploads above 4MB, use create-upload-session.',
      },
    ],
    opts.registeredTools
  );
  if (opts.readOnly) parts.push('This server is read-only; write operations are disabled.');
  if (opts.multiAccount)
    parts.push('Multiple accounts: pass the account parameter when required (see list-accounts).');
  if (!opts.orgMode)
    parts.push('Work/school-only tools require starting the server with --org-mode.');
  return parts.join(' ');
}

const DISCOVERY_MODE_INSTRUCTIONS_ADDON =
  'DISCOVERY MODE ADD-ON: Graph is reached via search-tools → get-tool-schema → execute-tool (plus auth helpers). ' +
  'Workflow: (1) call search-tools with short natural-language keywords (BM25-ranked); ' +
  '(2) call get-tool-schema(tool_name) to see the parameters, required fields, and enum values; ' +
  '(3) call execute-tool with tool_name exactly as returned and parameters shaped per the schema. ' +
  'Skipping get-tool-schema is the leading cause of Graph 400 errors here. ' +
  'If search-tools returns no matches, retry with shorter or different keywords.';

/**
 * Full MCP `initialize.instructions` string: general guidance for every mode, plus a discovery-only suffix when applicable.
 */
export function buildMcpServerInstructions(
  opts: McpInstructionsContext & { discovery: boolean }
): string {
  const general = buildGeneralMcpInstructions(opts);
  if (!opts.discovery) return general;
  return `${general} ${DISCOVERY_MODE_INSTRUCTIONS_ADDON}`;
}
