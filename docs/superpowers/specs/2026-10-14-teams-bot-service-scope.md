# Scope: app-only (unattended) Teams MCP service

Research-only scope, 2026-10-14. Nothing implemented. Everything marked **verified** was read from the current Microsoft Learn page or the public Graph docs source today; everything marked **unverified** was not, and must be tested before it is relied on.

## 1. Problem and goals

The ms-365 fork (`/Users/maxs/gitrepos/ms-365-mcp-server`) does Teams as the signed-in person: a delegated token from a device-code login, cached in MSAL, refreshed silently. Every message it sends is from Max or Olga, and it stops working when the refresh token dies. The agents in the `hermes-family` stack need a Teams presence that:

- runs unattended, with no human login to renew;
- is scoped to the teams an owner chose, not the whole tenant;
- posts under its own identity (a bot), so nothing it says is attributed to a person;
- co-deploys next to the m365 servers the way `docglean-mcp` does, as its own container with its own bearer token.

Decision already taken: a separate project, not a mode of the upstream server. This document scopes that project.

Non-goals are listed in section 5.

## 2. Verified capability table

Sources read today:

- Send message in channel (v1.0 and beta): https://learn.microsoft.com/en-us/graph/api/chatmessage-post?view=graph-rest-1.0 and the generated permission tables `api-reference/{v1.0,beta}/includes/permissions/chatmessage-post-permissions.md`, `channel-post-messages-permissions.md`, `chat-post-messages-permissions.md` in github.com/microsoftgraph/microsoft-graph-docs-contrib.
- Reply in channel: https://learn.microsoft.com/en-us/graph/api/chatmessage-post-replies?view=graph-rest-1.0
- Reactions: https://learn.microsoft.com/en-us/graph/api/chatmessage-setreaction?view=graph-rest-1.0
- Edit: https://learn.microsoft.com/en-us/graph/api/chatmessage-update?view=graph-rest-1.0
- List channel messages: https://learn.microsoft.com/en-us/graph/api/channel-list-messages?view=graph-rest-1.0 ; replies, chat messages, team members, channels, tabs, chat members: the matching `*-permissions.md` includes in the docs repo.
- Protected APIs: `concepts/teams-protected-apis.md` (github.com/microsoft/microsoft-graph-docs-1 mirror; the Learn URL https://learn.microsoft.com/en-us/graph/teams-protected-apis returned 404 to the fetcher today, so the mirror text is what was read).
- RSC permissions: https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/rsc/resource-specific-consent and https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/rsc/grant-resource-specific-consent ; permission descriptions from `concepts/permissions-reference.md` (RSC table).
- Bot Framework: https://learn.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-connector-api-reference?view=azure-bot-service-4.0 (Conversation operations), https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/send-proactive-messages , https://learn.microsoft.com/en-us/microsoftteams/platform/teams-sdk/in-depth-guides/message-reactions , https://learn.microsoft.com/en-us/azure/bot-service/bot-service-quickstart-registration?view=azure-bot-service-4.0

Legend: Delegated = what our m365 fork uses today. App-only = client-credentials token, tenant-wide permission consented by an admin in Entra. RSC = client-credentials token, `.Group`/`.Chat` permission granted per team or chat by installing a Teams app. Bot = Bot Framework Connector REST (Azure Bot + Teams app with a bot).

| Tool family (fork tool names)                                                                                                                                             | Delegated                                                  | App-only Graph (tenant-wide)                                                                                                                          | RSC Graph (per team/chat)                                                                                                                                                                                                                                                                                                                                                  | Bot Framework                                                                                                                                                                                                                               |
| ------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Read channel messages / replies / hosted content (`list-channel-messages`, `get-channel-message`, `list-channel-message-replies`, `list-channel-message-hosted-contents`) | `ChannelMessage.Read.All` **verified**                     | `ChannelMessage.Read.All` **verified**; this is a _protected API_ needing Microsoft's request form **verified**                                       | `ChannelMessage.Read.Group` is the documented least-privileged application permission on list messages and list replies **verified**. Whether a `.Group`-only app also needs the protected-API form is **unverified** (the form page says the listed APIs "all use application permissions" and does not exempt RSC; Inditex runs on `.Group` without mentioning the form) | No read-history operation exists in the Connector API (operations are create/send/reply/update/delete activity, members, attachments) **verified**                                                                                          |
| Send new channel post / reply (`send-channel-message`, `reply-to-channel-message`)                                                                                        | `ChannelMessage.Send` **verified**                         | `Teamwork.Migrate.All` only, and "Application permissions are only supported for migration" **verified** (same on beta). Not usable for live channels | `ChannelMessage.Send.Group` exists in the permissions reference ("Allows the app to send messages to this team's channels, without a signed-in user") **verified**, but it does not appear in the send or reply API permission tables on v1.0 or beta **verified**. Whether Graph POST honours it is **unverified**; assume no                                             | `SendToConversation` (new thread) and `ReplyToActivity` / conversation id `channelId;messageid=threadId` (reply in thread) **verified**; app with the bot must be installed in the team **verified**; message is from the bot, not a person |
| Edit message (`update-channel-message`, `update-channel-message-reply`, `update-chat-message`)                                                                            | `ChannelMessage.ReadWrite` / `Chat.ReadWrite` **verified** | Only the `policyViolation` property, needs a Communications DLP licence **verified**. Useless for us                                                  | none listed **verified**                                                                                                                                                                                                                                                                                                                                                   | `UpdateActivity` on the bot's own messages **verified**                                                                                                                                                                                     |
| Delete message                                                                                                                                                            | not fetched                                                | not fetched                                                                                                                                           | not fetched                                                                                                                                                                                                                                                                                                                                                                | `DeleteActivity` on the bot's own messages **verified**                                                                                                                                                                                     |
| Reactions (`set-/unset-channel-message-reaction`, `set-/unset-chat-message-reaction`)                                                                                     | `ChannelMessage.Send` / `Chat.ReadWrite` **verified**      | "Not supported." **verified**                                                                                                                         | none **verified**                                                                                                                                                                                                                                                                                                                                                          | Legacy Bot Framework: inbound `messageReaction` events only; the Teams SDK v2 (TS/C#/Python) has `api.conversations.addReaction/deleteReaction`, any message in a conversation the agent is in, limit two reactions per second **verified** |
| Chats: list messages (`list-chat-messages`, `get-chat-message`)                                                                                                           | `Chat.Read` **verified**                                   | `Chat.Read.All` (protected API) **verified**                                                                                                          | `ChatMessage.Read.Chat` is the least-privileged application permission **verified**; requires the app installed in that chat                                                                                                                                                                                                                                               | none                                                                                                                                                                                                                                        |
| Chats: send (`send-chat-message`, `reply-to-chat-message`)                                                                                                                | `ChatMessage.Send` **verified**                            | `Teamwork.Migrate.All` only (migration) **verified**                                                                                                  | none                                                                                                                                                                                                                                                                                                                                                                       | Proactive message to a group chat the app is installed in, or 1:1 if installed for the user **verified**                                                                                                                                    |
| Chats: list (`list-chats` is `/me/chats`)                                                                                                                                 | `Chat.ReadBasic`                                           | `/users/{id}/chats` with `Chat.Read.All`-class permission: **not fetched**                                                                            | not applicable                                                                                                                                                                                                                                                                                                                                                             | `GetConversations` lists conversations the bot is in **verified**                                                                                                                                                                           |
| Members (`list-team-members`, `list-chat-members`)                                                                                                                        | `TeamMember.Read.All` / `Chat.ReadBasic` **verified**      | `TeamMember.Read.All` / `ChatMember.Read.All` **verified**                                                                                            | `TeamMember.Read.Group` / `ChatMember.Read.Chat` **verified**                                                                                                                                                                                                                                                                                                              | `GetConversationPagedMembers` **verified** (what Inditex uses)                                                                                                                                                                              |
| Channels (`list-team-channels`, `get-team-channel`)                                                                                                                       | `Channel.ReadBasic.All` **verified**                       | `Channel.ReadBasic.All` etc. **verified**                                                                                                             | `ChannelSettings.Read.Group` **verified**                                                                                                                                                                                                                                                                                                                                  | not applicable                                                                                                                                                                                                                              |
| Tabs (`list-channel-tabs`)                                                                                                                                                | `TeamsTab.Read.All` **verified**                           | `TeamsTab.Read.All` **verified**                                                                                                                      | `TeamsTab.Read.Group` **verified**                                                                                                                                                                                                                                                                                                                                         | not applicable                                                                                                                                                                                                                              |
| Team/channel create, update, delete, add/remove member                                                                                                                    | various `*.ReadWrite.All`                                  | tenant-wide                                                                                                                                           | `Channel.Create.Group`, `Channel.Delete.Group`, `ChannelSettings.ReadWrite.Group`, `TeamSettings.ReadWrite.Group` exist **verified**; no member-write RSC                                                                                                                                                                                                                  | not applicable                                                                                                                                                                                                                              |

Two oddities worth knowing:

- `ChannelMember.Read.Group` is in Inditex's manifest, but the RSC page lists it as supported for neither delegated nor application mode **verified**. Harmless, probably useless; leave it out.
- The Graph send page's "Permissions for chat" table repeats `ChannelMessage.Send` where `ChatMessage.Send` is expected; the generated include for `chat-post-messages` says `ChatMessage.Send`. Documentation slip, not a capability difference.

**The single most important finding:** Microsoft Graph cannot post a live channel or chat message as an application. The only application permission on the send and reply APIs is `Teamwork.Migrate.All`, documented as migration-only, on both v1.0 and beta. An unattended Teams writer therefore has to be a Bot Framework bot, which is exactly why Inditex carries an Azure Bot. Reads, members, channels and tabs do work app-only, and have per-team RSC variants.

## 3. How Inditex actually does it

Source read: `src/mcp_teams_server/{__init__,teams,config}.py`, `app/manifest.json`, `doc/MS-Teams-setup.md` (copies under `scratchpad/inditex/`).

- Two clients from one Entra app + client secret (`TEAMS_APP_ID`, `TEAMS_APP_PASSWORD`, `TEAMS_APP_TENANT_ID`): a Graph `GraphServiceClient(ClientSecretCredential, scopes=[".default"])` and an Agents SDK `CloudAdapter(MsalConnectionManager)` for the Bot Connector (`__init__.py` lines 69-86).
- **Reads are Graph app-only**: `read_threads` and `read_thread_replies` call `/teams/{TEAM_ID}/channels/{TEAMS_CHANNEL_ID}/messages[/{id}/replies]` with `$top`, and validate the `@odata.nextLink` cursor against the configured team/channel path (`teams.py` lines 98-124, 386-502). The setup doc says to prefer `ChannelMessage.Read.Group` over tenant-wide `ChannelMessage.Read.All`.
- **Writes are Bot Framework, not Graph**: `start_thread` builds a synthetic `conversationUpdate` continuation activity against a hard-coded `service_url` (`https://smba.trafficmanager.net/emea/`), calls `adapter.continue_conversation`, then `send_activities` with `conversation=<channel>` to open a new thread; `update_thread` does `conversations.send_to_conversation(conversation_id=f"{channel};messageid={thread_id}")` to reply in a thread (`teams.py` lines 134-161, 210-360). Mentions are `<at>` entities resolved from the member list. The sender is `ChannelAccount(id=app_id, name="MCP Bot")`.
- **Members are Bot Framework** too: `get_conversation_paged_member` on the Teams connector client (`teams.py` lines 182-208).
- Nothing listens: no messaging endpoint is deployed ("you will not need to deploy any webapp into Azure"). The bot only ever calls out.
- Fixed scope: one `TEAM_ID` and one `TEAMS_CHANNEL_ID` from env; no allowlist beyond that, no audit log, no bearer auth on its HTTP transport.

What must exist in the tenant for that to work (setup doc, grant-RSC page, Azure Bot page):

1. Entra app registration with a client secret (single-tenant; multi-tenant bot creation was deprecated after 31 July 2025 **verified**).
2. An **Azure Bot** resource bound to that app id, with the Teams channel enabled. This needs an Azure subscription (free F0 tier exists). No messaging endpoint is needed when the bot only sends.
3. A **Teams app package** (manifest 1.19 in Inditex's case) with `bots[]` pointing at the app id and `authorization.permissions.resourceSpecific` listing the `.Group` application permissions. The RSC doc says `webApplicationInfo.id` must be the Entra app id; Inditex's manifest omits `webApplicationInfo` and relies on `bots[].botId`. Whether that is enough for RSC is **unverified**; include `webApplicationInfo` to be safe.
4. The app **installed into each team** by a team owner (or member, per tenant setting); installing is what grants the RSC permissions. The grant is visible at `GET /beta/teams/{groupId}/permissionGrants` **verified**.

Implications: every message is visibly from the bot; the bot can only act in teams it is installed in; a human still has to click "install" per team; and the Azure Bot adds an Azure dependency the m365 server and docglean do not have.

## 4. What this repo already has that is non-delegated

Read: `src/auth.ts`, `src/obo-client.ts`, `src/oauth-provider.ts`, `src/cli.ts`, `src/secrets.ts`, `src/server.ts` (grep), README auth sections.

- `src/auth.ts`: `PublicClientApplication` only. Device code, interactive browser, silent refresh, multi-account cache, cache coherency plugin. Delegated throughout.
- `src/obo-client.ts`: `ConfidentialClientApplication.acquireTokenOnBehalfOf` for `--obo` in HTTP mode. Needs an incoming user token; still delegated.
- `src/oauth-provider.ts`: MCP `ProxyOAuthServerProvider` that verifies a caller's bearer against `/me`. Delegated.
- `MS365_MCP_OAUTH_TOKEN` and `--trust-proxy-auth`: bearer pass-through; whatever token the caller supplies.
- `src/secrets.ts`: env or Azure Key Vault for client id, tenant id, client secret. The secret is used only for OBO and the OAuth token/refresh endpoints.
- **There is no client-credentials path.** No `acquireTokenByClientCredential` anywhere in `src/`. `server.ts` even bails on `--obo` without `--http`.

Reusable by analogy (patterns, not imports, since it is a different package): the secrets provider shape (`secrets.ts`), the structured audit log contract (`src/audit-log.ts`: tool, principal, target resource id, outcome, never bodies), response scrubbing ideas (`src/response-scrubbing.ts`), the two-listener/network-isolation thinking from the attachment proxy docs, and the e2e doc style (`docs/e2e-attachment-to-markdown.md`). Not reusable: the whole MSAL cache machinery, the OAuth proxy, `endpoints.json` tool generation (the new service has a handful of hand-written tools).

## 5. Shape by analogy with docglean

docglean (`/Users/maxs/gitrepos/docglean-mcp`, private repo `msoukhomlinov/docglean-mcp`): Python 3.12, `src/docglean_mcp/` package, `pyproject.toml`, `pydantic-settings` config with a `DOCGLEAN_MCP_` env prefix, streamable HTTP MCP at `/mcp`, refuses to bind `0.0.0.0` without `DOCGLEAN_MCP_BEARER_TOKEN`, `/healthz` and `/readyz`, two-stage Dockerfile, non-root uid 10001, pip stripped from the runtime layer, `HEALTHCHECK` on `/healthz`, `SECURITY.md` with a written threat model, GHCR publish on `v*.*.*` tags, squash merges, specs under `docs/superpowers/specs/`. Co-deployed in the `hermes-family` compose stack next to `m365-max-mcp` and `m365-olga-mcp`; the m365 README (around lines 1020-1050) shows the compose pairing with a shared bearer and per-network isolation.

Proposed shape for this project, same mould:

- **Own repo**: `msoukhomlinov/teams-bot-mcp` (name open), private like docglean.
- **Language**: TypeScript on Node 22, to match the fork and because the Teams SDK v2 for TypeScript (`@microsoft/teams.apps` / `@microsoft/teams.api`) is the one place reactions are documented as an outbound API. Python with the Agents SDK (what Inditex uses) is the alternative; see open question 3.
- **Container**: one image, non-root, `HEALTHCHECK`, config by env with a `TEAMS_MCP_` prefix, secret via env or a mounted file (`_FILE` variant), bearer token required on `/mcp` when bound off-host.
- **Transport**: streamable HTTP at `/mcp`, stdio optional for local testing.
- **Co-deployment**: one more service in the hermes-family compose stack; it needs egress to `graph.microsoft.com`, `login.microsoftonline.com` and `smba.trafficmanager.net`, and no link to the m365 containers at all (they share nothing).
- **How an agent chooses**: the servers have disjoint tool names and distinct `initialize.instructions`. The m365 server says "you are acting as <person>"; this one says "you are the <bot name> bot in teams X, Y; messages you post are visibly from the bot; use the m365 server when something must be sent as the person". Keep this server's tool names prefixed (`teams-bot-*` or similar) so nothing collides with the fork's `send-channel-message`.

## 6. Architecture options

**A. Graph app-only + RSC for reads, Bot Framework for writes (the Inditex model).** Full read and write unattended, per-team scope, bot identity. Needs Entra app, Azure Bot, Teams app with bot + RSC, install per team. Reactions via Teams SDK v2 connector API. Most capable; most setup; Azure subscription required.

**B. Read-only Graph app-only + RSC.** Entra app + Teams app with only RSC application permissions (no `bots[]`), installed per team. No Azure Bot, no Azure subscription, no connector. Tools: channels, messages, replies, members, hosted content. Writes stay with the delegated m365 server (as the person, via its `teams-write` preset). Smallest build and smallest blast radius. Does not give the agents an identity of their own.

**C. Write-only bot (no Graph at all).** Azure Bot + Teams app with a bot; post, reply, edit own, react, list members, all via the connector; nothing can be read back except what the bot itself sent. Cheap but blind; only sensible as the second half of A.

**D. Tenant-wide application permissions (`ChannelMessage.Read.All`, `Chat.Read.All`).** Rejected: protected-API request form, reads every channel in the tenant, no per-team consent, and still cannot write.

**E. Fork InditexTech/mcp-teams-server (Apache-2.0) and harden it.** Fastest way to A: it already works. Costs: Python (fine, docglean is Python), hard-coded single team/channel, EMEA service URL, no bearer auth, no audit log, Bot Framework "hacks" the authors flag themselves, and a licence header to preserve.

**Recommendation: A, built in two phases, with B as phase 1.** Phase 1 ships the read side (B) and, critically, settles the one empirical unknown (does `ChannelMessage.Read.Group` work without the protected-API form in our tenant) before any bot work. Phase 2 adds the Azure Bot, the `bots[]` manifest entry and the write tools. Design the manifest, config and allowlist for A from day one so phase 2 is an addition, not a rewrite. Note that adding a bot to an already-installed Teams app means a manifest version bump and an app update in each team.

Why not B alone: the user's stated reason for this project is unattended agents with their own presence; posting as Max from an agent is what the fork already does and is the thing to move away from. Why not E: the per-team allowlist, bearer auth, audit and secret handling are the parts that matter for us and are the parts Inditex lacks; rewriting those around their code is not less work than writing six tools next to a known-good reference.

## 7. MVP tool list

Phase 1 (read, Graph + RSC):

| Tool                                               | Graph call                                                                                | RSC permission               |
| -------------------------------------------------- | ----------------------------------------------------------------------------------------- | ---------------------------- |
| `teams-bot-list-teams`                             | from config only (the allowlist); no Graph                                                | none                         |
| `teams-bot-list-channels`                          | `GET /teams/{team}/channels`                                                              | `ChannelSettings.Read.Group` |
| `teams-bot-list-threads`                           | `GET /teams/{team}/channels/{ch}/messages?$top=N` with validated `@odata.nextLink` cursor | `ChannelMessage.Read.Group`  |
| `teams-bot-read-thread`                            | `GET .../messages/{id}` plus `/replies` paged                                             | `ChannelMessage.Read.Group`  |
| `teams-bot-list-members` / `teams-bot-find-member` | `GET /teams/{team}/members`                                                               | `TeamMember.Read.Group`      |
| `teams-bot-get-hosted-content`                     | `GET .../hostedContents/{id}/$value` (images in messages), optional                       | `ChannelMessage.Read.Group`  |

Phase 2 (write, Bot Connector via Teams SDK v2 or Agents SDK):

| Tool                                    | Connector operation                                                         |
| --------------------------------------- | --------------------------------------------------------------------------- |
| `teams-bot-post`                        | `SendToConversation` to the channel conversation; returns the new thread id |
| `teams-bot-reply`                       | `ReplyToActivity` / send to `channelId;messageid=threadId`                  |
| `teams-bot-edit`                        | `UpdateActivity`, own messages only                                         |
| `teams-bot-delete`                      | `DeleteActivity`, own messages only                                         |
| `teams-bot-react` / `teams-bot-unreact` | Teams SDK `addReaction` / `deleteReaction`                                  |

Mentions: resolve `@Name` to a member id from `list-members` and emit an `<at>` entity, as Inditex does. The fork's mention-shorthand spec (`docs/superpowers/specs`, recent commits on `integration/hermes-family`) is the UX to mirror.

Explicit non-goals for v1: 1:1 and group chats (needs chat-scope RSC and per-chat install; revisit), tenant-wide permissions, team or channel management, tabs, change notifications or any inbound webhook, message import/migration, sending as a person, private and shared channels until tested, multi-tenant.

## 8. Security model

- **Least privilege by construction**: only `.Group` application RSC permissions in the manifest (`ChannelMessage.Read.Group`, `ChannelSettings.Read.Group`, `TeamMember.Read.Group`; phase 2 adds nothing on the Graph side because writes do not go through Graph). No tenant-wide permission is ever consented in Entra, so there is nothing to approve on the protected-API form and the app cannot read a team it is not installed in. RSC "permissions aren't attributed to a user... the app can be allowed to perform actions that the user can't" (grant-RSC page), which is why the next point exists.
- **Service-side allowlist, enforced not advisory**: `TEAMS_MCP_ALLOWED_TEAMS` (team id to optional channel-id list). Every tool resolves team and channel against it before any network call, and pagination cursors are validated against the configured path like Inditex's `_validate_pagination_cursor`. Being installed in a team is not enough to be reachable from the tools; being in the allowlist is not enough without the install. Both must hold.
- **Caller auth**: bearer token on `/mcp`, required when bound off-host, same rule as docglean. One token per agent container if more than one agent uses it, so the audit log can name the caller.
- **Secrets**: the Entra client secret (or certificate) read from a file (`_FILE`) or env; never a tool argument; never logged; the process drops it from the environment after reading it if a worker model is ever introduced. Prefer a **certificate** over a secret if both the MSAL client-credentials path and the bot connector auth accept one in the chosen SDK; **unverified** for the Agents/Teams SDK auth managers, verify in the spike. Rotation is an Entra operation plus a container restart.
- **If the secret leaks**: the holder can, for every team the app is installed in, read all channel messages and members, and (phase 2) post, edit own, delete own and react as the bot. They cannot read any other team, any chat, any mailbox, or act as a user, and cannot escalate through Graph because no tenant-wide permission is consented. Containment: rotate the secret in Entra, uninstall the Teams app from the affected teams (removes the grants), and check Entra sign-in logs for the service principal. Consider a Conditional Access policy for workload identities to pin the service principal to the host's egress IP.
- **Audit log**: structured JSON like `src/audit-log.ts` in the fork: timestamp, caller (bearer id), tool, team id, channel id, thread or message id, outcome and error class; never message bodies. Reads are audited as well as writes.
- **Content hygiene**: responses carry message HTML from arbitrary users; strip to text, keep hosted-content references as Graph paths, and treat every body as data (prompt-injection surface, same posture as the mail tools).
- **Rate limits**: Graph Teams throttling plus the documented two reactions per second; honour `Retry-After`.

## 9. Setup steps and who does them

Tenant admin (Entra and Teams admin roles):

1. Entra app registration in the **same tenant** as the teams (sideloading an RSC app registered elsewhere fails with `WebApplicationInfoIdOfSideloadedAppMustBeInTheSameTenantAsUser` unless the installer is a tenant admin **verified**). Single tenant. Add a client secret or upload a certificate. No API permissions need to be added in Entra for RSC ("You don't need to add the non-RSC permissions to the app manifest as Microsoft Entra admin center stores them", and RSC ones live only in the Teams manifest **verified**).
2. Leave the tenant RSC setting at its default `ManagedByMicrosoft` (allows team RSC consent), or set `Set-MgBetaTeamRscConfiguration -State EnabledForAllApps` **verified**. Check it is not `DisabledForAllApps`.
3. Allow the app in Teams: either enable custom app upload for the installing user (Teams admin centre app setup policy) or publish the package to the org app catalogue, and make sure app permission policies do not block it.
4. Phase 2 only: create the **Azure Bot** resource (Azure subscription, F0) bound to the existing app id, single-tenant, enable the Microsoft Teams channel. No messaging endpoint.
5. Not needed unless option D is chosen: the protected-API form at https://aka.ms/teamsgraph/requestaccess.

Team owner:

6. Install the Teams app into the team. The install dialog shows the RSC permissions; accepting grants them for that team only.
7. Give the operator the team id (groupId) and the channel ids from "Get link to channel".

Operator (us):

8. Put team and channel ids in `TEAMS_MCP_ALLOWED_TEAMS`, mount the secret, start the container, run the probe tool (`verify` that lists channels and the last message of each allowed channel) and confirm `GET /beta/teams/{groupId}/permissionGrants` shows our `clientAppId`.

## 10. Risks

1. **RSC reads gated by the protected-API programme.** The protected-API page lists "List channel messages" and "List replies" with no RSC exemption; community answers and Inditex's working setup say `.Group` does not need the form. If it does, phase 1 stalls on a Microsoft review (weekly). Mitigation: task 0 below is a half-day probe with nothing but an app registration and a sideloaded manifest; do it before writing code.
2. **Graph write via `ChannelMessage.Send.Group` is unverified** and probably not honoured. Plan on the bot; test the Graph path once in the probe because it would remove the Azure Bot dependency if it worked.
3. **Service URL region.** Inditex hard-codes `smba.trafficmanager.net/emea/`; the proactive-messaging page gives the global `https://smba.trafficmanager.net/teams/` and warns not to hard-code. With no inbound endpoint there is no incoming activity to copy it from. Needs a test against our (Australian) tenant.
4. **Azure dependency** for phase 2: an Azure subscription, an Azure Bot resource, and the Microsoft Teams channel on it. Multi-tenant bots are deprecated, so the registration must be single-tenant.
5. **Private and shared channels** may behave differently under RSC (a Microsoft Q&A answer says change notifications are blocked for them; reads are unknown). Test; exclude from v1 if they fail.
6. **Identity expectations**: everything the agents post is visibly from the bot. If stakeholders expect posts "from Max", this project is the wrong tool and the delegated server is the right one.
7. **Two Teams SDK generations** (Bot Framework / Agents SDK versus Teams SDK v2). Reactions are only in the newer one. Pick one in the spike; mixing them is how Inditex ended up with turn-state hacks.
8. **Install friction per team**: a human owner clicks install, and every manifest change (adding the bot in phase 2) is an app update per team.

## 11. Open questions for the user

1. Is an Azure subscription available in the target tenant for the Azure Bot resource, and is the Azure dependency acceptable at all? (Decides whether phase 2 exists.)
2. Is read-only (B) enough for the first release, or do the agents need to post unattended from day one?
3. Language: TypeScript (matches the fork; Teams SDK v2 has reactions) or Python (matches docglean; fork or copy Inditex)?
4. Which tenant and team is the pilot (yames.com test tenant versus intellectit.com.au), and are you tenant admin there, or does someone else have to do steps 1-5?
5. What should the bot be called, and should its posts carry a footer (the fork has a signoff feature; the same question applies here)?
6. Are group or 1:1 chats wanted later? That changes the manifest (chat-scope RSC) and the install story (per chat).
7. Client secret or certificate for the app credential?
8. Repo name and whether it stays private like docglean.

## 12. Rough task breakdown for a later plan

0. **Probe (no code, half a day, user-driven)**: register the app, build a manifest with only the three read RSC permissions, sideload into a test team, get a client-credentials token with curl, call list channels, list messages, list replies, list members; also try one `POST .../messages` with `ChannelMessage.Send.Group` in the manifest to settle risk 2. Record status codes only. This answers risks 1 and 2 and open question 4.
1. Repo scaffold from docglean conventions: package layout, config with env prefix, bearer auth, `/healthz`, Dockerfile (non-root, no package manager in runtime), `SECURITY.md`, CI, GHCR publish on tags.
2. Auth: MSAL client credentials (`.default`), secret or certificate from file or env, token cache in memory only.
3. Allowlist and path validation module with its own tests (team, channel, cursor).
4. Phase 1 tools (section 7) with a thin Graph client: pagination, `Retry-After`, HTML-to-text, hosted-content references.
5. Audit log and MCP `initialize.instructions` text (who the bot is, which teams, when to use the m365 server instead).
6. Compose integration in hermes-family: service, networks, secret mount, agent config; e2e doc in the style of `docs/e2e-attachment-to-markdown.md`.
7. Phase 2: Azure Bot, manifest bump with `bots[]`, connector client, post/reply/edit/delete, service-URL discovery, mention entities.
8. Phase 2b: reactions via Teams SDK v2 if that SDK was chosen; otherwise document as unsupported.
9. Hardening pass: secret-leak drill (rotate, uninstall, confirm 401/403), rate-limit behaviour, private-channel test, Teams app manifest validation in CI.

---

## Decisions (user, 2026-10-14)

- **Writes from day one.** Phase 1 is no longer read-only: the Azure Bot and the Teams app with the bot installed in the team are part of the first release. Posts appear as the bot.
- **Language: TypeScript.**
- **Pilot tenant: intellectit.com.au** (the real tenant). Admin consent and a team owner's agreement are needed before the Phase 0 probe.
- **Credential: client secret.** Mitigate with a short expiry, a documented rotation step, a secret file rather than an env value in the container, and the service-side team/channel allowlist, since a leaked secret works from anywhere until it expires or is rotated.

## Still open

- Is an Azure subscription available for the Azure Bot resource? Writes from day one depend on it.
- Is the user tenant admin on intellectit.com.au, or does someone else grant consent?
- Bot display name, and whether its posts carry a footer or signoff.
- Group or 1:1 chats later? That changes the manifest and per-chat install.
- Repo name, and private like docglean?
- The two unverified Graph points (protected-API form for `.Group`-only apps; Graph POST with `ChannelMessage.Send.Group`) are settled only by the Phase 0 probe.

Status: scoping draft produced by a research subagent; its Microsoft Learn citations have not been independently re-checked. No design approval yet, so no repo or tenant changes have been made.
