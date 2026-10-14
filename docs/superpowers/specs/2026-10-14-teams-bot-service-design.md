# Teams bot service: design

A separate, private TypeScript MCP service that gives agents unattended, per-team Microsoft Teams access under their own bot identity. It is built outside the upstream `ms-365-mcp-server`, in the same way as `docglean-mcp`, and is co-deployed next to the m365 servers.

Related documents, all in this folder:

- `2026-10-14-teams-bot-service-scope.md`: research and the capability table (what Graph can and cannot do as an application). Parts of it are superseded by this design; see the note at its end.
- `2026-10-14-teams-bot-phase0-probe.md`: the no-code tenant probe that settles the conditional items below and records the response fixtures the tests use.

## Why a separate service

The fork's Teams tools act as the signed-in person. Agents that run unattended need their own identity, access limited to chosen teams, and no human login to renew. Microsoft Graph cannot post a live channel or chat message as an application (the only application permission on the send and reply APIs is `Teamwork.Migrate.All`, documented as migration-only), and a bot can only edit or delete its own Connector messages. So unattended writes go through a Bot Framework bot, and reads use Graph app-only with per-team Resource Specific Consent (RSC) permissions.

## Decisions

Made with the user on 2026-10-14:

- Writes ship in the first release, not later.
- TypeScript.
- Pilot tenant: intellectit.com.au.
- Credential: client secret.
- Graph and the Bot Connector are called through thin REST clients (`fetch` plus client-credentials tokens). No Teams SDK and no Microsoft 365 Agents SDK.
- The team and channel allowlist is a mounted JSON config file with aliases and a per-channel write flag.

Changed after the independent QA review (see Review log): mentions use the fork's inline shorthand; no list-channels tool; no Graph-POST path; a single global `serviceUrl`.

## Goals

- Agents read and write in chosen Teams channels with no human sign-in.
- Access is bounded twice: by the RSC install per team, and by the service's own allowlist per channel, and neither can be bypassed by a crafted id or cursor.
- Every action is audited without message bodies.
- A leaked secret reaches only the teams where the app is installed.

## Non-goals (v1)

- 1:1 and group chats, which need chat-scope RSC and a per-chat install.
- Tenant-wide application permissions (`ChannelMessage.Read.All`, `Chat.Read.All`). Rejected: protected-API request form, every channel in the tenant, no per-team consent.
- Team or channel management, tabs, change notifications, any inbound webhook, message import or migration.
- Acting as a person (the fork's delegated tools do that).
- Listing channels from Graph, hosted-content (inline image) download, reactions, and a per-channel write budget. Each can be added later without changing the design.
- Private and shared channels until tested, multi-tenant use, per-agent bearer tokens.

## Architecture

One container, one process, streamable HTTP at `/mcp`. The image runs as a non-root user. The service only calls out; it deploys no Bot Framework messaging endpoint (see Conditional on the probe).

### Components

| Component          | Responsibility                                                                                                                                         | Depends on             |
| ------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------ | ---------------------- |
| Config loader      | Reads and validates the config file and the secret and token files at startup; fails fast on any problem                                               | filesystem             |
| Allowlist resolver | Resolves `team` and `channel` aliases to ids, and enforces the per-channel `write` flag                                                                | config                 |
| Token provider     | Client-credentials tokens for the Graph scope and the Bot Connector scope; cached until about 5 minutes before expiry; one in-flight request per scope | config, `fetch`        |
| Graph client       | Thin `fetch` GETs to `https://graph.microsoft.com/v1.0/` only; validates cursors; retries reads                                                        | token provider         |
| Connector client   | Thin `fetch` calls to the configured `serviceUrl` for post, reply, edit, delete                                                                        | token provider, config |
| Mention converter  | Rewrites the inline mention shorthand into Teams text and mention entities                                                                             | none                   |
| Audit log          | One JSON line per tool call to stderr; never bodies                                                                                                    | none                   |
| HTTP layer         | `/mcp`, `/healthz`, `/readyz`; bearer and `Origin` checks                                                                                              | config                 |
| Tools              | The MCP tool layer; each tool validates input, resolves aliases, calls a client, audits                                                                | all of the above       |

### Token authority

Both tokens are fetched with the client-credentials flow from `https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token`, never from the `botframework.com` authority (the bot is single-tenant):

- Graph: scope `https://graph.microsoft.com/.default`.
- Bot Connector: scope `https://api.botframework.com/.default`.

A failure at the token endpoint (for example `invalid_client` for a wrong or expired secret) is its own error class, `token_endpoint`, in logs and in tool errors. Rotating the secret means restarting the container.

### Config file

Path from `TEAMS_BOT_CONFIG`. JSON:

```json
{
  "tenantId": "<tenant guid>",
  "appId": "<app guid>",
  "secretFile": "/run/secrets/teams_bot_secret",
  "tokenFile": "/run/secrets/teams_bot_token",
  "serviceUrl": "https://smba.trafficmanager.net/teams/",
  "teams": [
    {
      "alias": "ops",
      "id": "<team group id guid>",
      "channels": [
        { "alias": "alerts", "id": "19:abc@thread.tacv2", "write": true },
        { "alias": "general", "id": "19:def@thread.tacv2", "write": false }
      ]
    }
  ]
}
```

`serviceUrl` is optional and defaults to the global URL above. `tokenFile` holds the bearer token the HTTP layer requires. Other environment: `TEAMS_BOT_HOST`, `TEAMS_BOT_PORT`, and `TEAMS_BOT_ALLOWED_ORIGINS` (default empty).

Startup validation fails with a clear message when:

- a required field is missing, an unknown key is present, or the secret or token file is empty or unreadable;
- `tenantId`, `appId` or a team `id` is not a GUID;
- a channel `id` does not match `^19:[^/]+@thread\.(tacv2|skype)$`;
- an alias is duplicated within its scope, or the same channel id appears more than once anywhere in the file;
- `serviceUrl` is not `https` with host exactly `smba.trafficmanager.net`.

### Input trust rules

The agent is untrusted: it can be steered by text in channel messages. These rules are what keep the allowlist real.

- **Ids from the agent.** `team` and `channel` are aliases resolved through config; the agent never supplies a raw team or channel id. `thread` and `message` must match `^\d{1,20}$` (Teams message ids are numeric). Anything else is a validation error and nothing is sent.
- **Encoding.** Every path segment is built with `encodeURIComponent`, including the whole Connector conversation id (`<channel>;messageid=<thread>`). Table tests cover `..`, `/`, `;messageid=` and `%2F` injection.
- **Graph host.** The Graph client builds URLs itself on `https://graph.microsoft.com/v1.0/` and attaches the Graph token to nothing else.
- **Cursors.** `cursor` is a Graph `@odata.nextLink` that was returned earlier. It is accepted only when all of the following hold: scheme `https`; host exactly `graph.microsoft.com`; no userinfo, port or fragment; the decoded path is exactly the expected path for the tool (`/v1.0/teams/{team}/channels/{channel}/messages` for `list-threads`, `.../messages/{thread}/replies` for `read-thread`); no `..`, backslash or `%2F` in it. Anything else is rejected and no request is made.
- **Connector host.** The Connector token goes only to the configured `serviceUrl`, never to a URL taken from a message or a tool argument.

### Tools

All tool names are prefixed `teams-bot-`. `team` is optional when exactly one team is configured. A channel is always given by its alias.

| Tool                     | Input                                                              | Output                                                                                                           |
| ------------------------ | ------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------- |
| `teams-bot-list-teams`   | none                                                               | allowed teams and channels with aliases and write flags, from config only                                        |
| `teams-bot-list-threads` | `channel`, `team?`, `top?` (1-50, default 20), `cursor?`           | `{ threads: [message], nextCursor? }`                                                                            |
| `teams-bot-read-thread`  | `channel`, `thread`, `team?`, `top?` (1-50, default 20), `cursor?` | `{ root?: message, replies: [message], nextCursor? }`                                                            |
| `teams-bot-list-members` | `team?`, `name?`                                                   | `[{ id, userId, displayName, email, roles }]`; `name` is a case-insensitive substring over displayName and email |
| `teams-bot-post`         | `channel`, `text`, `team?`                                         | `{ threadId, messageId }` (the same id)                                                                          |
| `teams-bot-reply`        | `channel`, `thread`, `text`, `team?`                               | `{ messageId }`                                                                                                  |
| `teams-bot-edit`         | `channel`, `thread`, `message`, `text`, `team?`                    | `{ messageId }`                                                                                                  |
| `teams-bot-delete`       | `channel`, `thread`, `message`, `team?`                            | `{ deleted: true }`                                                                                              |

Read output and ordering:

- **Message shape.** `{ id, createdDateTime, lastModifiedDateTime, from: { kind, id, displayName }, isOwn, text }`. `kind` is `user`, `bot`, `app` or `system`; `isOwn` is true when the author is this app (`application.id` equals `appId`). Messages whose `messageType` is not `message`, and deleted messages, are left out.
- **`text`** is the message HTML reduced to plain text with one small existing dependency (`html-to-text`), not a hand-written parser. Hosted-content references are kept as Graph paths and not downloaded.
- **Ordering.** Graph orders channel messages by the last modified time of the whole reply chain, not by creation. `list-threads` returns them in Graph's order; both timestamps are in the output so an agent can tell.
- **`read-thread` paging.** `root` appears only on the first page. `top` applies to replies. Later pages carry replies only.
- **Untrusted text.** The tool descriptions label `text` and `displayName` as untrusted data.

Write rules:

- `post`, `reply`, `edit` and `delete` require the channel's `write` flag.
- Edit and delete apply to the bot's own messages only; that is a Bot Connector limit, and the tool descriptions say so. An agent can use `isOwn` to tell which messages qualify.
- Text longer than a fixed constant (80,000 characters, Microsoft's recommended size) is rejected before sending, and a 413 `MessageSizeTooBig` from the Connector maps to the same error.

### Mentions

Mentions use the same inline shorthand as the fork's Teams tools: `<at user-id="<entra-object-id>">Display Name</at>` inside `text`. The agent gets `userId` from `teams-bot-list-members`. Rules:

- `user-id` must be a GUID. Any `<at ... user-id` tag that is not exactly `<at user-id="<guid>">Name</at>` (single quotes, extra attributes, nested markup) is a tool error and nothing is sent.
- Each valid tag becomes `<at>Name</at>` in the activity text plus a mention entity whose `mentioned.id` is the `userId`. The name is XML-escaped.
- `post`, `reply` and `edit` all accept the shorthand. An edit that omits it drops the mentions the old text had, and the tool description says so.
- The exact activity text format and entity shape are confirmed by the probe before they are fixed in code.

### Data flow

- **Read:** tool, then input validation, then alias resolution, then Graph with the RSC token, then HTML reduced to text, then audit, then result.
- **Write:** tool, then input validation, then alias and `write` check, then mention conversion, then Connector call, then audit, then the new id.

### Manifest

The production Teams app package contains:

- the latest stable Teams app manifest schema at implementation time (the probe records the version that uploads);
- `webApplicationInfo.id` equal to `appId`;
- one entry in `bots[]` with `botId` equal to `appId`, scope `team`, and `isNotificationOnly` true;
- RSC application permissions exactly `ChannelMessage.Read.Group` and `TeamMember.Read.Group`. Nothing else; in particular not `ChannelMessage.Send.Group`, `ChannelSettings.Read.Group`, `ChannelMember.Read.Group`, `Member.Read.Group` or `Owner.Read.Group`, which the Inditex manifest carries.

Side effect to document for users: the bot has no endpoint, so anyone who messages or @mentions it gets no reply.

## Security model

- **Two boundaries.** RSC limits the app to teams where an owner installed it. The service allowlist limits it to the listed channels and enforces read-only channels. The input trust rules above stop an id or cursor from crossing either boundary.
- **Leaked secret.** The holder can, for every team where the app is installed, read channel messages and members, and post, edit and delete as the bot. They cannot read other teams, chats or mailboxes, and cannot act as a user. Containment: rotate the secret in Entra, restart the container, and uninstall the Teams app from the affected teams.
- **Secret handling.** The secret and the bearer token come from files, never from argv or logs. A short expiry and a documented rotation step are part of the setup.
- **Bearer and origin.** The bearer token is always required, compared in constant time. Any request that carries an `Origin` header is rejected unless the origin is listed in `TEAMS_BOT_ALLOWED_ORIGINS`, which closes browser DNS-rebinding access. There is no loopback exemption.
- **Untrusted content.** Message bodies are data, never instructions, the same posture as the m365 mail tools. Accepted residual risk: text read from a read-only channel can still steer an agent into posting it in a write channel elsewhere. The allowlist does not stop that, and the tool descriptions say so.
- **Audit.** One JSON line per call: time, tool, team, channel, thread and message id, outcome (`ok`, `error` or `unknown`), error class. Reads are logged as well as writes; bodies are not. There is no `caller` field because there is a single bearer token; per-agent tokens are a later change.
- **Errors.** Graph and Connector errors surface as tool errors with the error code and a short message, never the raw response, never any token or secret.

## Error handling and retries

- **Validation errors** (unknown alias, bad id, bad cursor, write to a read-only channel, bad mention tag, over-length text): tool error, nothing sent.
- **Reads** (Graph): retried up to 3 attempts on 429, 502, 503 and 504, honouring `Retry-After`, then a tool error that says the service is throttled or unavailable.
- **Writes** (Connector): retried on 429 only. A 5xx or a timeout may still have delivered the message, so the result is an error that says the outcome is unknown and to check `teams-bot-list-threads` before retrying, and the audit outcome is `unknown`. This avoids posting twice.
- **401 and 403:** a tool error naming the likely cause (secret expired, app not installed in that team) without exposing credentials.
- **Token endpoint failures:** class `token_endpoint`, see Token authority.

## Operability

- `GET /healthz`: liveness, returns 200 and nothing else.
- `GET /readyz`: fetches both tokens (result cached 30 seconds) and returns 200 or 503 with no detail in the body. A wrong or expired secret shows here, not on the first tool call.
- The MCP `initialize` response carries instructions saying what this server is and when to use the m365 server instead: this server acts as the bot in the configured teams; use the m365 server when something must be sent as a person. That is how an agent chooses between the two.
- A short secret expiry has no built-in warning in v1. The setup documentation lists the expiry date next to the rotation steps.

## Testing

- Vitest, with an injected `fetch`. No network in the default run.
- Table-driven allowlist and config cases: unknown alias, single-team default, read-only channel, duplicate aliases and channel ids, unknown keys, bad channel id pattern, bad `serviceUrl` host.
- Input trust cases: `thread` and `message` traversal and `;messageid=` injection, cursor with a foreign host, wrong path, `beta` path, `hostedContents` path, encoded separators.
- Mention conversion: valid tag, several tags, bad GUID, single quotes, extra attributes, nested markup, XML escaping of the name, edit without the shorthand.
- Retry policy: reads retry and honour `Retry-After`; a write 5xx or timeout yields `unknown` and is not retried; 429 on a write is retried.
- Read output: `isOwn`, `kind`, system and deleted messages left out, `read-thread` paging.
- Error mapping with no secret in any message; `/readyz` failing on a bad secret.
- Request and response fixtures come from the redacted bodies recorded in the Phase 0 probe.
- One opt-in live test, behind an environment variable, against the probe team.

## Conditional on the probe

| Unknown                                                                                                                                             | Isolated behind   | Fallback if the probe says no                                                                  |
| --------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------- | ---------------------------------------------------------------------------------------------- |
| Bot registration route (Q1): the Developer Portal may fail with a multitenant-deprecation error                                                     | Setup docs only   | Azure Bot resource, which needs an Azure subscription; or drop unattended writes               |
| RSC-only reads without the protected-API form (Q2)                                                                                                  | Graph client      | Reconsider read paths; the service cannot ship reads as designed                               |
| Outbound-only Connector with the global `serviceUrl` and no messaging endpoint, new-thread body, reply, edit and delete conversation-id shapes (Q3) | Connector client  | A small inbound endpoint so the service can learn the `serviceUrl` and conversation references |
| Graph message ids equal Connector activity ids, so a human-started thread can be replied to by its Graph id (Q4)                                    | Connector client  | Keep a mapping from Graph ids to activity ids, or limit replies to threads the bot started     |
| Mention text format and entity shape                                                                                                                | Mention converter | Whatever shape the probe shows works                                                           |

If the Connector needs an inbound endpoint, this design changes in one place: the deployment gains an HTTPS listener for Bot Framework traffic.

## Open items

These need the user and are not blockers for the spec:

- **Azure subscription.** Community reports say the Developer Portal route now fails, so writes from day one probably need an Azure Bot resource and an Azure subscription. Answer this before the probe.
- Whether the user is tenant admin on intellectit.com.au or another admin grants consent.
- The bot display name, and whether its posts carry a footer or signoff.
- The repository name, and whether it is private like docglean (assumed private).
- Whether group or 1:1 chats are wanted later.

## Success criteria

With the probe passed and a configured team:

1. An agent lists, reads and pages through an allowed channel and lists its members, with no human sign-in.
2. An agent posts, replies, edits and deletes as the bot in a channel with `write: true`, mentions members with the shorthand, and replies in a thread a human started.
3. Every request outside the allowlist, with a malformed id, cursor or mention tag, or to a `write: false` channel is refused with a clear error and nothing is sent.
4. A write that times out is reported as unknown and never retried.
5. The audit log shows each call without bodies, and no log line or error contains a secret.
6. `/readyz` reports a bad secret before the first tool call.
7. The default test run passes with no network.

## Process

- The service lives in a new private repository, created only after the user approves the name and visibility. It is not created by this spec.
- Same toolchain conventions as the fork (TypeScript, vitest, eslint, prettier), GHCR images on tags, a `SECURITY.md`, a non-root image, one `html-to-text` dependency.
- No AI attribution anywhere: commits, pull requests, code comments, documentation.

## Review log

An independent QA review (Opus) found two Critical issues, thirteen Important, and a set of Minor findings. Rulings:

- **Fixed:** cursor and token-host validation; numeric id validation and path encoding; Connector write retry policy; loopback bearer exemption removed and `Origin` checked; manifest defined; `/healthz` and `/readyz`; `token_endpoint` error class; `isOwn` and author `kind` on reads; message ordering and `read-thread` paging documented; config validation holes closed; edit and mentions interplay documented; fixed size constant and 413 mapping; token authority stated; HTML-to-text dependency named; residual cross-channel injection risk stated; scope document marked superseded where it differs.
- **Changed from the chat-approved design:** mentions use the fork's inline shorthand; `teams-bot-list-channels` removed (and with it `ChannelSettings.Read.Group`); the Graph-POST path and `ChannelMessage.Send.Group` removed; one global `serviceUrl` with a host check and no per-team override; `caller` dropped from audit.
- **Not adopted:** a per-channel write budget (optional; a later change).
