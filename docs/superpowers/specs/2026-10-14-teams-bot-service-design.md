# Teams bot service: design

A separate, private TypeScript MCP service that gives agents unattended, per-team Microsoft Teams access under their own bot identity. It is built outside the upstream `ms-365-mcp-server`, in the same way as `docglean-mcp`, and is co-deployed next to the m365 servers.

Related documents, all in this folder:

- `2026-10-14-teams-bot-service-scope.md`: research and the capability table (what Graph can and cannot do as an application).
- `2026-10-14-teams-bot-phase0-probe.md`: the no-code tenant probe that settles the conditional items below.

## Why a separate service

The fork's Teams tools act as the signed-in person. Agents that run unattended need their own identity, access limited to chosen teams, and no human login to renew. Microsoft Graph cannot post a live channel or chat message as an application (the only application permission on the send and reply APIs is `Teamwork.Migrate.All`, documented as migration-only), so unattended writes need a Bot Framework bot. Reads work app-only with per-team Resource Specific Consent (RSC) permissions.

## Decisions

Made with the user on 2026-10-14:

- Writes ship in the first release, not later.
- TypeScript.
- Pilot tenant: intellectit.com.au.
- Credential: client secret.
- Graph and the Bot Connector are called through thin REST clients (`fetch` plus client-credentials tokens). No Teams SDK and no Microsoft 365 Agents SDK.
- The team and channel allowlist is a mounted JSON config file with aliases and a per-channel write flag.

## Goals

- Agents read and write in chosen Teams channels with no human sign-in.
- Access is bounded twice: by the RSC install per team, and by the service's own allowlist per channel.
- Every action is audited without message bodies.
- A leaked secret reaches only the teams where the app is installed.

## Non-goals (v1)

- 1:1 and group chats, which need chat-scope RSC and a per-chat install.
- Tenant-wide application permissions (`ChannelMessage.Read.All`, `Chat.Read.All`). Rejected: protected-API request form, every channel in the tenant, no per-team consent.
- Team or channel management, tabs, change notifications, any inbound webhook, message import or migration.
- Acting as a person (the fork's delegated tools do that).
- Private and shared channels until tested, multi-tenant use.
- Hosted-content (inline image) download, reactions and unreactions unless the probe confirms the call shape. Both can be added after v1 without changing the design.

## Architecture

One container, one process, streamable HTTP at `/mcp`. A bearer token read from a file is required for any request that does not come from loopback. The image runs as a non-root user. The service only calls out; it deploys no Bot Framework messaging endpoint (see Conditional on the probe).

### Components

| Component | Responsibility | Depends on |
|---|---|---|
| Config loader | Reads and validates the config file and secret file at startup; fails fast on any problem | filesystem |
| Allowlist resolver | Resolves `team` and `channel` aliases to ids, and enforces the per-channel `write` flag | config |
| Token provider | Client-credentials tokens for the Graph scope and the Bot Connector scope; cached until about 5 minutes before expiry; one in-flight request per scope | config, `fetch` |
| Graph client | Thin `fetch` GETs; retries 429 and 503 honouring `Retry-After`, at most 3 attempts; validates `@odata.nextLink` cursors against the allowed team and channel path | token provider |
| Connector client | Thin `fetch` calls to the Bot Connector for post, reply, edit, delete | token provider, config (`serviceUrl`) |
| Mention resolver | Resolves names to member ids using the cached member list (5 minute TTL) | Graph client |
| Audit log | One JSON line per tool call to stderr: time, caller, tool, team, channel, message id, outcome, error class; never bodies | none |
| Tools | The MCP tool layer; each tool resolves aliases first, then calls a client, then audits | all of the above |

### Config file

Path from `TEAMS_BOT_CONFIG`. JSON:

```json
{
  "tenantId": "<tenant guid>",
  "appId": "<app guid>",
  "secretFile": "/run/secrets/teams_bot_secret",
  "serviceUrl": "https://smba.trafficmanager.net/emea/",
  "maxTextLength": 20000,
  "teams": [
    {
      "alias": "ops",
      "id": "<team group id guid>",
      "serviceUrl": "https://smba.trafficmanager.net/apac/",
      "channels": [
        { "alias": "alerts", "id": "19:...@thread.tacv2", "write": true },
        { "alias": "general", "id": "19:...@thread.tacv2", "write": false }
      ]
    }
  ]
}
```

Other environment: `TEAMS_BOT_TOKEN_FILE` (the bearer token file), `TEAMS_BOT_HOST` and `TEAMS_BOT_PORT`.

Startup validation fails with a clear message when:

- a required field is missing, or the secret or token file is empty or unreadable;
- `tenantId`, `appId` or a team `id` is not a GUID;
- an alias is duplicated within its scope;
- a `serviceUrl` is not an `https` URL;
- `maxTextLength` is present and not a positive integer.

A team-level `serviceUrl` overrides the top-level one. The connector token is only ever sent to a configured `serviceUrl`, never to a URL taken from a message or a tool argument.

### Tools

All tool names are prefixed `teams-bot-`. `team` is optional when exactly one team is configured. A channel is always given by its alias.

| Tool | Input | Output |
|---|---|---|
| `teams-bot-list-teams` | none | allowed teams and channels with aliases and write flags, from config only |
| `teams-bot-list-channels` | `team?` | the allowed channels for the team, with the display name and membership type from Graph |
| `teams-bot-list-threads` | `channel`, `team?`, `top?` (1-50, default 20), `cursor?` | `{ threads: [{ id, createdDateTime, from, text }], nextCursor? }` |
| `teams-bot-read-thread` | `channel`, `thread`, `team?`, `top?`, `cursor?` | the root message and its replies, same message shape, with `nextCursor?` |
| `teams-bot-list-members` | `team?`, `name?` | `[{ id, userId, displayName, email, roles }]`; `name` is a case-insensitive substring over displayName and email |
| `teams-bot-post` | `channel`, `text`, `team?`, `mentions?` | `{ threadId, messageId }` |
| `teams-bot-reply` | `channel`, `thread`, `text`, `team?`, `mentions?` | `{ messageId }` |
| `teams-bot-edit` | `channel`, `thread`, `message`, `text`, `team?` | `{ messageId }` |
| `teams-bot-delete` | `channel`, `thread`, `message`, `team?` | `{ deleted: true }` |

Notes:

- **Cursors.** `cursor` is the Graph `@odata.nextLink` returned as `nextCursor`. The Graph client accepts it only when its path is under the resolved team and channel.
- **Message shape.** `text` on reads is the message HTML reduced to plain text. Hosted-content references are kept as Graph paths and not downloaded.
- **Edit and delete** apply to the bot's own messages only; that is a Bot Connector limit, and the tool descriptions say so.
- **Write limits.** `post`, `reply`, `edit` and `delete` require the channel's `write` flag. Text longer than `maxTextLength` (optional, default 20000 characters) is rejected with a clear error.

### Mentions

`mentions` is an array of member names or ids. Each entry is resolved through the cached member list for the team:

- an exact id or a unique name match resolves;
- an ambiguous name is a tool error that lists the candidates and their `userId` values; it never guesses;
- no match is a tool error.

For each resolved mention, the first occurrence of `@<the string passed>` in `text` becomes `<at>Resolved Display Name</at>`; if the string does not occur in `text`, `<at>Resolved Display Name</at> ` is prepended. The activity carries a matching mention entity for each.

### Data flow

- **Read:** tool, then alias resolution, then Graph with the RSC token, then HTML reduced to text, then audit, then result.
- **Write:** tool, then alias and `write` check, then mention resolution, then Connector call, then audit, then the new thread or message id.

## Security model

- **Two boundaries.** RSC limits the app to teams where an owner installed it. The service allowlist limits it to the listed channels and enforces read-only channels.
- **Leaked secret.** The holder can, for every team where the app is installed, read channel messages and members, and post, edit and delete as the bot. They cannot read other teams, chats or mailboxes, and cannot act as a user. Containment: rotate the secret in Entra and uninstall the Teams app from the affected teams.
- **Secret handling.** The secret and the bearer token come from files, never from argv or logs. A short expiry and a documented rotation step are part of the setup.
- **Untrusted content.** Message bodies are data, never instructions, the same posture as the m365 mail tools. They are reduced to text and returned inside a structured result.
- **Bearer token.** Compared in constant time; requests from loopback may omit it, anything else may not.
- **Audit.** Reads and writes are logged; message bodies are not.
- **Errors.** Graph and Connector errors surface as tool errors with the error code and a short message, never the raw response, and never any token or secret.

## Error handling

- Unknown alias, write to a read-only channel, ambiguous or missing mention, over-length text: tool error, nothing sent.
- 429 and 503 from Graph or the Connector: retried up to 3 attempts honouring `Retry-After`, then a tool error that says the service is throttled.
- 401 or 403: a tool error naming the likely cause (secret expired, app not installed in that team) without exposing credentials.
- A cursor outside the resolved team and channel: tool error, request not made.

## Testing

- Vitest, with an injected `fetch`. No network in the default run.
- Table-driven allowlist cases: unknown alias, single-team default, read-only channel, duplicate aliases at startup.
- Cursor validation, mention resolution and ambiguity, text over the limit, retry and `Retry-After` handling, error mapping with no secret in any message.
- Request shapes for the Connector and Graph are taken from the responses recorded in the Phase 0 probe.
- One opt-in live test, behind an environment variable, against the probe team.

## Conditional on the Phase 0 probe

| Unknown | Isolated behind | Fallback if the probe says no |
|---|---|---|
| Bot registration without Azure (Q1) | Setup docs only, no code | Azure Bot resource and subscription, or the service-account option |
| RSC-only reads without the protected-API form (Q2) | Graph client | Reconsider read paths; the service cannot ship reads as designed |
| Graph POST with `ChannelMessage.Send.Group` (Q3) | Connector client | Post through Graph, and reply too if the same holds for replies; the bot remains for edit and delete |
| `serviceUrl` source and region (Q4) | Config `serviceUrl` fields | A small inbound endpoint so the service can learn the URL from `conversationUpdate` |
| Reaction call shape | Connector client | Leave reactions out of v1 |
| RSC-only Teams app accepted without `bots[]` | Setup docs only | Include a placeholder tab or the bot entry |

If the probe shows the connector token cannot reach a channel without an inbound endpoint, this design changes in one place: the deployment gains an HTTPS listener for Bot Framework traffic, and the config gains no new fields beyond a public URL.

## Open items

These need the user and are not blockers for the spec:

- The bot display name, and whether its posts carry a footer or signoff.
- The repository name, and whether it is private like docglean (assumed private).
- Whether group or 1:1 chats are wanted later.
- Whether an Azure subscription is available, if Q1 comes back "no".
- Whether the user is tenant admin on intellectit.com.au or another admin grants consent.

## Success criteria

With the probe passed and a configured team:

1. An agent lists, reads and pages through an allowed channel and lists its members, with no human sign-in.
2. An agent posts, replies, edits and deletes as the bot in a channel with `write: true`, and mentions members by name.
3. Every request outside the allowlist, to a `write: false` channel, or with an ambiguous mention is refused with a clear error and nothing is sent.
4. The audit log shows each call without bodies, and no log line or error contains a secret.
5. The default test run passes with no network.

## Process

- The service lives in a new private repository, created only after the user approves the name and visibility. It is not created by this spec.
- Same toolchain conventions as the fork (TypeScript, vitest, eslint, prettier), GHCR images on tags, a `SECURITY.md`, a non-root image.
- No AI attribution anywhere: commits, pull requests, code comments, documentation.
