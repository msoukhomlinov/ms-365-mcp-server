# Teams bot service: Phase 0 probe checklist

No-code probe in the **intellectit.com.au** tenant. It settles the unknowns in `2026-10-14-teams-bot-service-design.md` (section "Conditional on the probe") before any code, and it records the redacted response bodies the service's tests use as fixtures. Half a day, one tenant admin, one throwaway team.

Portal labels below are from memory and may differ; when a step does not match the portal, record what you saw instead. "Expect" lines are expectations from the research, not verified facts. The point of the probe is to replace them with recorded results.

## Ground rules

- **Throwaway team only.** Create `m365-bot-probe` with one channel `probe`. Nothing in this probe touches a production team, chat or mailbox.
- **Secrets never go in chat, commits or shell history.** Read them with `read -s APP_SECRET`. Give the secret a 1-day expiry.
- **Record verbatim.** For every call, paste the HTTP status and the full body into the results table or a fixture file. Do not paraphrase errors.
- **Fixtures stay out of git.** Save response bodies under a folder outside any repository (for example `~/probe-fixtures/`), and redact bearer tokens, tenant ids and real user names before they are ever copied into the service repo.
- **Clean up at the end** (last section), even if a step fails.

## The questions

| #   | Question                                                                                                                                                                                                                       | Why it matters                                                  |
| --- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------- |
| Q1  | Which bot registration route works: Azure Bot resource, or the Teams Developer Portal?                                                                                                                                         | Decides whether an Azure subscription is required               |
| Q2  | Does an app holding only per-team RSC permissions read channel messages **without** the protected-API request form?                                                                                                            | Decides whether reads work without a Microsoft approval request |
| Q3  | Does the Bot Connector accept new-thread, reply, edit, delete and mention posts with the **global `serviceUrl`** and **no messaging endpoint**? Which conversation ids work for edit and delete of a root post versus a reply? | Decides whether the outbound-only design works                  |
| Q4  | Do Graph message ids equal Connector activity ids, so the bot can reply to a thread a human started using the id Graph returns?                                                                                                | Decides whether agents can reply to existing threads            |
| Q5  | Is a team without the app installed refused?                                                                                                                                                                                   | Proves the per-team boundary                                    |

## Before you start

Answer the Azure subscription question. Community reports say the Developer Portal route now fails with a multitenant-deprecation error, so writes from day one probably need an Azure Bot resource, which needs a subscription. If no subscription is available, run only Steps 3 and 4 (reads) and stop: unattended writes are then not possible.

## Setup (tenant admin)

1. Create the team `m365-bot-probe` and channel `probe`. Note the **team id** (groupId) and **channel id** (`19:...@thread.tacv2`) from "Get link to channel". Export them as `TEAM_ID` and `CHANNEL_ID`.
2. Check the tenant RSC setting is not disabled: `Get-MgBetaTeamRscConfiguration` should report `ManagedByMicrosoft` or `EnabledForAllApps`, not `DisabledForAllApps`.
3. In the Teams admin centre, make sure custom app upload is allowed for you (or plan to upload as admin).
4. Post two top-level messages by hand in `probe`, and reply once to the first. Note the id of the first message (the **human thread**). You will need more than one message so a page of one forces a `nextLink`.

## Step 1: register the bot (Q1)

1. **Default route:** create an Entra app registration (single tenant) with a client secret, then an Azure Bot resource bound to that app id, single-tenant, with the Microsoft Teams channel enabled and **no messaging endpoint**. Record: subscription used, tier, whether the portal accepted an empty endpoint.
2. **Optional alternative, time-boxed to 15 minutes:** in the Teams Developer Portal, **Tools > Bot management > New bot**. Record whether it asked for an Azure subscription, whether it created an app id and secret, and the verbatim error if it failed.
3. Record the app id as `APP_ID`, the tenant id as `TENANT`, and the secret via `read -s APP_SECRET`.

**Q1 result:** route used, and whether the other route works: **\_\_**

## Step 2: build and install the test Teams app

Create a manifest with the latest stable Teams app manifest schema that uploads, and record its version. Include:

- `webApplicationInfo.id` = `APP_ID`
- one `bots[]` entry with `botId` = `APP_ID`, scope `team`, and `isNotificationOnly: true`
- `authorization.permissions.resourceSpecific` with exactly the application permissions `ChannelMessage.Read.Group` and `TeamMember.Read.Group`. Do **not** add `ChannelMessage.Send.Group`, `ChannelSettings.Read.Group` or any other permission; production does not hold them.

Upload it, then install it into `m365-bot-probe` as the team owner. Record the permission list the install dialog showed.

Get a Graph token and confirm the grant exists:

```bash
read -s APP_SECRET   # paste the secret, press enter
TOKEN=$(curl -s -X POST "https://login.microsoftonline.com/$TENANT/oauth2/v2.0/token" \
  -d "client_id=$APP_ID" --data-urlencode "client_secret=$APP_SECRET" \
  -d "scope=https://graph.microsoft.com/.default" -d "grant_type=client_credentials" | jq -r .access_token)
curl -s -H "Authorization: Bearer $TOKEN" "https://graph.microsoft.com/beta/teams/$TEAM_ID/permissionGrants" | jq .
```

Expect: a grant row with `clientAppId` equal to `APP_ID`. Also record the token's roles, which should show no tenant-wide Teams permissions (expectation, unverified):

```bash
echo "$TOKEN" | jq -R 'split(".")[1] | gsub("-";"+") | gsub("_";"/") | @base64d | fromjson | .roles'
```

## Step 3: Graph reads with RSC only (Q2, Q5)

```bash
G="https://graph.microsoft.com/v1.0"; H="Authorization: Bearer $TOKEN"
curl -s -w '\nHTTP %{http_code}\n' -H "$H" "$G/teams/$TEAM_ID/members"
curl -s -w '\nHTTP %{http_code}\n' -H "$H" "$G/teams/$TEAM_ID/channels/$CHANNEL_ID/messages?\$top=1"
```

The second call should return one message and an `@odata.nextLink`. Then fetch the next page by calling the `nextLink` value exactly as returned, then read a single message and the replies of the human thread:

```bash
curl -s -w '\nHTTP %{http_code}\n' -H "$H" "<the @odata.nextLink value>"
curl -s -w '\nHTTP %{http_code}\n' -H "$H" "$G/teams/$TEAM_ID/channels/$CHANNEL_ID/messages/$HUMAN_THREAD_ID"
curl -s -w '\nHTTP %{http_code}\n' -H "$H" "$G/teams/$TEAM_ID/channels/$CHANNEL_ID/messages/$HUMAN_THREAD_ID/replies?\$top=1"
```

Save each response body as a fixture. Do not request protected-API access for any of this.

- **Q2 pass:** the messages calls return 200 with data. **Q2 fail:** a 4xx; paste the body, it may name the protected-API requirement.
- **Q5 negative control:** call the members and messages endpoints for a team where the app is **not** installed (any other test team id you own). Expect a refusal; record status and body.

## Step 4: bot posting with no messaging endpoint (Q3, Q4)

Get a Bot Connector token. The authority is your tenant, and the scope is different from Graph's:

```bash
BTOKEN=$(curl -s -X POST "https://login.microsoftonline.com/$TENANT/oauth2/v2.0/token" \
  -d "client_id=$APP_ID" --data-urlencode "client_secret=$APP_SECRET" \
  -d "scope=https://api.botframework.com/.default" -d "grant_type=client_credentials" | jq -r .access_token)
SVC="https://smba.trafficmanager.net/teams/"   # the global URL; try this first
C='Content-Type: application/json'; B="Authorization: Bearer $BTOKEN"
```

Run each call and record the status and full body. Save every body as a fixture.

**4a. New thread** (the documented body, with tenant fields). Expect a response carrying a conversation id and an activity id:

```bash
curl -s -w '\nHTTP %{http_code}\n' -X POST -H "$B" -H "$C" \
  -d "{\"isGroup\":true,\"bot\":{\"id\":\"28:$APP_ID\"},\"tenantId\":\"$TENANT\",\"channelData\":{\"channel\":{\"id\":\"$CHANNEL_ID\"},\"tenant\":{\"id\":\"$TENANT\"}},\"activity\":{\"type\":\"message\",\"text\":\"bot probe: new thread\"}}" \
  "${SVC}v3/conversations"
```

If the global URL fails, record the error, then try the regional URLs (`amer`, `emea`, `apac`) and record which work. Set `ROOT_ID` to the activity id returned and `CONV` to the returned conversation id.

**4b. Reply** into that thread:

```bash
curl -s -w '\nHTTP %{http_code}\n' -X POST -H "$B" -H "$C" \
  -d '{"type":"message","text":"bot probe: reply"}' \
  "${SVC}v3/conversations/$(jq -rn --arg c "$CHANNEL_ID;messageid=$ROOT_ID" '$c|@uri')/activities"
```

Record the activity id returned as `REPLY_ID`. Check in Teams that both messages show the bot as author.

**4c. Mention.** Post a reply whose text contains a mention of one real member, using that member's `userId` (Entra object id) from the Step 3 members response. Try the documented form: `text` containing `<at>Name</at>` plus an `entities` array with a `mention` entity whose `mentioned` is `{ "id": "<userId>", "name": "Name" }` and whose `text` exactly equals `<at>Name</at>`. Record whether the person is notified and how the message renders, and the exact working body.

**4d. Edit and delete, root versus reply.** For each, try the bare channel id and `channel;messageid=ROOT_ID` as the conversation id, and record which works:

```bash
# edit the bot's reply
curl -s -w '\nHTTP %{http_code}\n' -X PUT -H "$B" -H "$C" -d '{"type":"message","text":"bot probe: reply edited"}' \
  "${SVC}v3/conversations/<conversation id>/activities/$REPLY_ID"
# edit the bot's root post, then delete the reply, then delete the root post
curl -s -w '\nHTTP %{http_code}\n' -X DELETE -H "$B" "${SVC}v3/conversations/<conversation id>/activities/<activity id>"
```

Record for each of the four operations (edit reply, edit root, delete reply, delete root): the conversation id form that worked and the status. Also try editing the **human** message and record the refusal.

**4e. Reply to a human-started thread using the Graph id (Q4).** Use `HUMAN_THREAD_ID` (the id Graph returned in Step 3) as the `messageid`:

```bash
curl -s -w '\nHTTP %{http_code}\n' -X POST -H "$B" -H "$C" -d '{"type":"message","text":"bot probe: reply to human thread"}' \
  "${SVC}v3/conversations/$(jq -rn --arg c "$CHANNEL_ID;messageid=$HUMAN_THREAD_ID" '$c|@uri')/activities"
```

Record whether it lands in the human thread. Then confirm with Graph that the bot's new reply is visible in `.../messages/$HUMAN_THREAD_ID/replies`, and that the id Graph reports for the bot's own root post from 4a equals `ROOT_ID`.

**Findings to record about `serviceUrl`:** a bot with no messaging endpoint never receives a `conversationUpdate`, so it never sees the real service URL. Note whether the global `https://smba.trafficmanager.net/teams/` is reliable for this tenant or whether a region-specific URL was needed.

## Results table

| Q                                            | Result (pass / fail / partial) | HTTP status | Verbatim error or note |
| -------------------------------------------- | ------------------------------ | ----------- | ---------------------- |
| Q1 Azure Bot resource route                  |                                |             |                        |
| Q1 Developer Portal route                    |                                |             |                        |
| Q2 RSC-only reads without protected-API form |                                |             |                        |
| Q2 nextLink fetch and replies                |                                |             |                        |
| Q3 new thread, global serviceUrl             |                                |             |                        |
| Q3 reply                                     |                                |             |                        |
| Q3 mention post                              |                                |             |                        |
| Q3 edit reply, conversation id form          |                                |             |                        |
| Q3 edit root, conversation id form           |                                |             |                        |
| Q3 delete reply                              |                                |             |                        |
| Q3 delete root                               |                                |             |                        |
| Q3 edit of a human message refused           |                                |             |                        |
| Q4 reply to human thread by Graph id         |                                |             |                        |
| Q4 Graph root id equals activity id          |                                |             |                        |
| Q5 uninstalled team refused                  |                                |             |                        |
| Teams manifest schema version that uploaded  |                                |             |                        |

## Decisions the results drive

| Result                                                     | What it means for the design                                                                                                                      |
| ---------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| Q1 Azure Bot resource works                                | Writes from day one need an Azure subscription; keep it as the documented setup.                                                                  |
| Q1 neither route works                                     | Unattended writes are not possible; fall back to reads only, or the service-account option.                                                       |
| Q1 Developer Portal works                                  | Drop the Azure subscription requirement from the setup docs.                                                                                      |
| Q2 pass                                                    | Reads work with per-team consent only. No Microsoft access request.                                                                               |
| Q2 fail                                                    | Reads need the protected-API request, which is slow and not guaranteed. Reconsider read paths.                                                    |
| Q3 pass on the global URL                                  | The outbound-only design works. Fix the recorded body shapes in the Connector client.                                                             |
| Q3 needs a regional URL                                    | Keep one top-level `serviceUrl` and set it per deployment; update the default and the docs.                                                       |
| Q3 fails everywhere                                        | The service needs a small inbound endpoint so it can learn the `serviceUrl` and conversation references. That changes deployment (inbound HTTPS). |
| Q3 edit or delete only works with one conversation id form | Hard-code that form in the Connector client and its tests.                                                                                        |
| Q4 pass                                                    | Agents can reply to human-started threads using the Graph id.                                                                                     |
| Q4 fail                                                    | Limit replies to threads the bot started, or keep an id mapping.                                                                                  |
| Q5 pass                                                    | The per-team boundary is real. The service allowlist is defence in depth.                                                                         |
| Q5 fail                                                    | The app can read teams it is not installed in. Stop: the security model is wrong.                                                                 |

## Clean-up

1. Delete the secret from the app registration and `unset APP_SECRET TOKEN BTOKEN`.
2. Uninstall the Teams app from `m365-bot-probe`, delete the team, and delete the test app registration and any bot or Azure Bot resource.
3. Keep the filled-in results table and the redacted fixtures; they become part of the design and the test suite.
