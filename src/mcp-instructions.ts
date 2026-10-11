/** Shared context for MCP `initialize.instructions` (hosts that forward it to the model). */
export type McpInstructionsContext = {
  orgMode: boolean;
  readOnly: boolean;
  multiAccount: boolean;
  /**
   * The tool names this server actually registered (`resolveRegisteredToolNames`), not the flags
   * that shaped them: `--attachment-proxy` being on does not prove read-document exists, because
   * an --enabled-tools filter can drop it. A clause that names tools is emitted only where all of
   * them are here. Omitted means the tools the non-proxy wording names, i.e. the existing text.
   */
  registeredTools?: ReadonlySet<string>;
};

/**
 * One piece of guidance and the tools it directs the model to. Declaring the names next to the
 * text means no prose has to be classified: a sentence about a tool is worth saying exactly where
 * that tool can be called. `tools` must all be registered; `anyOf` needs at least one.
 */
type GuidanceClause = { text: string; tools?: readonly string[]; anyOf?: readonly string[] };

const DEFAULT_REGISTERED_TOOLS: ReadonlySet<string> = new Set([
  'list-users',
  'get-download-url',
  'download-bytes',
  'download-bytes-to-file',
  'upload-file-content',
  'create-upload-session',
]);

function emit(clauses: readonly GuidanceClause[], registered: ReadonlySet<string>): string[] {
  return clauses
    .filter(
      (clause) =>
        (clause.tools ?? []).every((tool) => registered.has(tool)) &&
        (!clause.anyOf || clause.anyOf.some((tool) => registered.has(tool)))
    )
    .map((clause) => clause.text)
    .filter((text) => text.trim().length > 0);
}

/**
 * The byte/document paragraph, one self-contained sentence per tool, so any combination of
 * registered tools yields true text. Without the proxy and with every byte tool registered this
 * is the upstream wording. Under --attachment-proxy download-bytes and get-download-url are not
 * registered and read-document takes their place. When none of them is registered the paragraph
 * is omitted rather than describing what is missing: an absence is not this module's to assert.
 */
function buildByteContentInstructions(registered: ReadonlySet<string>): string {
  const [lead, ...rest] = emit(
    [
      {
        tools: ['get-download-url'],
        text: 'for large drive/SharePoint file content, prefer get-download-url to resolve a pre-authenticated URL for out-of-band download.',
      },
      {
        tools: ['download-bytes'],
        text: 'Use download-bytes for authenticated byte reads such as mail attachments, profile photos, Teams hosted content, and meeting recordings.',
      },
      {
        tools: ['download-bytes', 'download-bytes-to-file'],
        text: 'In stdio mode (or HTTP with --http-local-file-tools), download-bytes-to-file writes those same authenticated bytes straight to a local absolute path instead of returning base64.',
      },
      {
        tools: ['get-download-url', 'download-bytes-to-file'],
        text: 'For large mail attachments and meeting recordings that is one out-of-band option; get-download-url is the other, on a server running with --enable-attachment-urls.',
      },
      {
        anyOf: ['download-bytes', 'get-download-url'],
        text: 'These tools take relative Microsoft Graph paths, not absolute URLs.',
      },
      {
        tools: ['read-document'],
        text:
          'Read documents with read-document, which returns markdown and takes a relative Microsoft ' +
          'Graph byte path: a mail or event attachment ' +
          '(/me/messages/{message-id}/attachments/{attachment-id}/$value; the /$value suffix is ' +
          'required), a raw message (/me/messages/{message-id}/$value), or a drive or SharePoint file ' +
          '(/drives/{drive-id}/items/{driveItem-id}/content). Absolute URLs are not accepted, and ' +
          'pages, offset and maxChars read a long document in parts. Byte payloads are stripped from ' +
          'tool results here: any contentBytes field, and any large base64 value. read-document ' +
          'mints a URL that is redeemed later with no Authorization header, so it refuses with ' +
          'identity_not_supported whenever Graph identity comes from the request (OAuth, OBO, or ' +
          "bearer mode) rather than from this server's own token cache.",
      },
    ],
    registered
  );
  if (lead === undefined) return '';
  return `Files / binary content: ${[lead, ...rest].join(' ')}`;
}

function buildGeneralMcpInstructions(opts: McpInstructionsContext): string {
  const registered = opts.registeredTools ?? DEFAULT_REGISTERED_TOOLS;
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
      { text: buildByteContentInstructions(registered) },
      {
        tools: ['upload-file-content', 'create-upload-session'],
        text: 'For uploads, upload-file-content takes a base64 string body (Graph allows 250MB, but the whole string passes through the agent context and a truncated one is written without error); use create-upload-session for anything but small files.',
      },
    ],
    registered
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
