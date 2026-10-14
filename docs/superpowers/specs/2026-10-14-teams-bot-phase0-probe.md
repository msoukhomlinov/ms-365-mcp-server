# Teams bot service: Phase 0 probe checklist

No-code probe in the **intellectit.com.au** tenant. It settles the unknowns listed in `2026-10-14-teams-bot-service-scope.md` before any code or Azure spend. Half a day, one tenant admin, one throwaway team.

Portal labels below are from memory and may differ; when a step does not match the portal, record what you saw instead. "Expect" lines are expectations from the scoping research, not verified facts. The point of the probe is to replace them with recorded results.

## Ground rules

- **Throwaway team only.** Create `m365-bot-probe` with one channel `probe`. Nothing in this probe touches a production team, chat or mailbox.
- **Secrets never go in chat, commits or shell history.** Read them with `read -s APP_SECRET`. Give the secret a 1-day expiry.
- **Record verbatim.** For every call, paste the HTTP status and the full error body into the results table. Do not paraphrase errors.
- **Clean up at the end** (last section), even if a step fails.

## The questions

| #   | Question                                                                                                                                       | Why it matters                                                          |
| --- | ---------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------- |
| Q1  | Can a bot be registered through the **Teams Developer Portal** with **no Azure subscription**?                                                 | Decides whether the Azure dependency exists at all                      |
| Q2  | Does an app holding only per-team RSC (`*.Group`) permissions read channel messages **without** the protected-API request form?                | Decides whether reads work without a Microsoft approval request         |
| Q3  | Does Graph `POST .../messages` honour `ChannelMessage.Send.Group` for an app-only caller?                                                      | If yes, top-level posts need no bot                                     |
| Q4  | Can the bot post a new thread and a reply through the Bot Connector with **no messaging endpoint**, and where does the `serviceUrl` come from? | Decides whether writes work and how the service learns the regional URL |
| Q5  | Does the RSC grant stay scoped: a team without the app installed is refused?                                                                   | Proves the per-team boundary                                            |

## Setup (tenant admin)

1. Create the team `m365-bot-probe` and channel `probe`. Note the **team id** (groupId) and **channel id** (`19:...@thread.tacv2`) from "Get link to channel". Export them as `TEAM_ID` and `CHANNEL_ID`.
2. Check the tenant RSC setting is not disabled: `Get-MgBetaTeamRscConfiguration` should report `ManagedByMicrosoft` or `EnabledForAllApps`, not `DisabledForAllApps`.
3. In the Teams admin centre, make sure custom app upload is allowed for you (or plan to upload as admin).

## Step 1: register the bot without Azure (Q1)

1. Open the Teams Developer Portal, **Tools > Bot management**, and choose **New bot**. Use the name `probe-bot` (the final display name is still open).
2. Record: did it ask for an Azure subscription? Did it create an app id and a client secret for you? Is the bot single-tenant? Record the app id as `APP_ID`, the tenant id as `TENANT`, and the secret via `read -s APP_SECRET`.
3. If the portal route is unavailable or demands an Azure subscription, stop this step and record it. Fall back to an Entra app registration (single tenant, client secret) and, only if Azure is acceptable, an Azure Bot resource with the Teams channel enabled and **no messaging endpoint**. Continue with that app id.

**Q1 result:** portal registration works without Azure: yes / no. Details: **\_\_**

## Step 2: build and install the test Teams app

Create a manifest (Developer Portal **Apps > New app**, or hand-written) with:

- `webApplicationInfo.id` = `APP_ID`
- `bots[]` with `botId` = `APP_ID`, scope `team`, `isNotificationOnly: true`
- `authorization.permissions.resourceSpecific` containing the application permissions `ChannelMessage.Read.Group`, `ChannelSettings.Read.Group`, `TeamMember.Read.Group` and `ChannelMessage.Send.Group`

Upload it, then install it into `m365-bot-probe` as the team owner. Record the permission list the install dialog showed.

Optional variant to answer a scoping doubt: repeat once with **no** `bots[]` and a placeholder static tab instead. Record whether an RSC-only app is accepted without a bot, since the scope doc assumed it would be.

Confirm the grant exists:

```bash
read -s APP_SECRET   # paste the secret, press enter
TOKEN=$(curl -s -X POST "https://login.microsoftonline.com/$TENANT/oauth2/v2.0/token" \
  -d "client_id=$APP_ID" --data-urlencode "client_secret=$APP_SECRET" \
  -d "scope=https://graph.microsoft.com/.default" -d "grant_type=client_credentials" | jq -r .access_token)
curl -s -H "Authorization: Bearer $TOKEN" "https://graph.microsoft.com/beta/teams/$TEAM_ID/permissionGrants" | jq .
```

Expect: a grant row with `clientAppId` equal to `APP_ID`. Also record the token's roles, which should show no tenant-wide Teams permissions (expectation, unverified):

```bash
echo "$TOKEN" | cut -d. -f2 | tr '_-' '/+' | base64 -D 2>/dev/null | jq '.roles'
```

## Step 3: reads with RSC only (Q2, Q5)

```bash
G="https://graph.microsoft.com/v1.0"; H="Authorization: Bearer $TOKEN"
curl -s -w '\nHTTP %{http_code}\n' -H "$H" "$G/teams/$TEAM_ID/channels"
curl -s -w '\nHTTP %{http_code}\n' -H "$H" "$G/teams/$TEAM_ID/members"
curl -s -w '\nHTTP %{http_code}\n' -H "$H" "$G/teams/$TEAM_ID/channels/$CHANNEL_ID/messages?\$top=5"
```

Post one message into `probe` by hand in the Teams client first, so there is something to read. Do not request protected-API access for this test.

- **Q2 pass:** the messages call returns 200 with the message. **Q2 fail:** a 4xx; paste the body, it may name the protected-API requirement.
- **Q5 negative control:** call the same three endpoints for a team where the app is **not** installed (use any other test team id you own). Expect a refusal; record status and body.

## Step 4: Graph POST with app-only (Q3)

```bash
curl -s -w '\nHTTP %{http_code}\n' -X POST -H "$H" -H "Content-Type: application/json" \
  -d '{"body":{"contentType":"text","content":"graph app-only post probe"}}' \
  "$G/teams/$TEAM_ID/channels/$CHANNEL_ID/messages"
```

Expect: refused, since the documented application permission for this call is migration-only. **Any success is a major finding:** it would make top-level posts possible without a bot. Record the exact status and body either way, and check the channel for a visible message.

## Step 5: bot posting with no messaging endpoint (Q4)

Get a Bot Connector token (note the different scope):

```bash
BTOKEN=$(curl -s -X POST "https://login.microsoftonline.com/$TENANT/oauth2/v2.0/token" \
  -d "client_id=$APP_ID" --data-urlencode "client_secret=$APP_SECRET" \
  -d "scope=https://api.botframework.com/.default" -d "grant_type=client_credentials" | jq -r .access_token)
```

The Inditex server hard-codes `https://smba.trafficmanager.net/emea/` as the service URL. Your tenant's region may differ. Try the value for the tenant's region first and record which service URLs work and which fail (candidates: `amer`, `emea`, `apac`):

```bash
SVC="https://smba.trafficmanager.net/apac/"
# new thread in the channel
curl -s -w '\nHTTP %{http_code}\n' -X POST -H "Authorization: Bearer $BTOKEN" -H "Content-Type: application/json" \
  -d "{\"isGroup\":true,\"bot\":{\"id\":\"28:$APP_ID\"},\"channelData\":{\"channel\":{\"id\":\"$CHANNEL_ID\"}},\"activity\":{\"type\":\"message\",\"text\":\"bot connector post probe\"}}" \
  "${SVC}v3/conversations"
```

The response should carry a conversation id of the form `<channel>;messageid=<threadId>`. Then reply into that thread:

```bash
CONV="<conversation id from the response>"
curl -s -w '\nHTTP %{http_code}\n' -X POST -H "Authorization: Bearer $BTOKEN" -H "Content-Type: application/json" \
  -d '{"type":"message","text":"bot connector reply probe"}' \
  "${SVC}v3/conversations/$CONV/activities"
```

Record: HTTP statuses, whether both messages appear in the channel, and that they show the bot as the author. This request shape is reconstructed from the Inditex source and the Bot Connector API, not run by us before, so if it is rejected, record the error and try the form the portal or SDK documentation shows.

**Service URL finding to record:** a bot with no messaging endpoint never receives a `conversationUpdate`, so it never sees the real `serviceUrl`. Note whether a fixed regional URL is reliable for this tenant, or whether the service needs a small endpoint or a configured URL per team.

## Results table

| Q                                            | Result (pass / fail / partial) | HTTP status | Verbatim error or note |
| -------------------------------------------- | ------------------------------ | ----------- | ---------------------- |
| Q1 portal registration without Azure         |                                |             |                        |
| Q2 RSC-only reads without protected-API form |                                |             |                        |
| Q3 Graph POST app-only                       |                                |             |                        |
| Q4 thread post via Bot Connector             |                                |             |                        |
| Q4 thread reply via Bot Connector            |                                |             |                        |
| Q4 working serviceUrl region                 |                                |             |                        |
| Q5 uninstalled team refused                  |                                |             |                        |
| RSC-only app accepted without `bots[]`       |                                |             |                        |

## Decisions the results drive

| Result                 | What it means for the design                                                                                                  |
| ---------------------- | ----------------------------------------------------------------------------------------------------------------------------- |
| Q1 yes                 | No Azure subscription needed. Remove that dependency from the scope.                                                          |
| Q1 no                  | Writes from day one require an Azure subscription and an Azure Bot resource, or the dedicated service-account option instead. |
| Q2 pass                | Reads work with per-team consent only. No Microsoft access request.                                                           |
| Q2 fail                | Reads need the protected-API request, which is slow and not guaranteed. Reconsider read paths.                                |
| Q3 pass                | Top-level posts can go through Graph. The bot would then only be needed for replies and edits.                                |
| Q3 fail                | As expected. Writes go through the bot.                                                                                       |
| Q4 pass                | The bot-only-outbound design works. Use the recorded `serviceUrl` approach.                                                   |
| Q4 fail on all regions | The service needs a messaging endpoint so it can learn the `serviceUrl`. That changes deployment (inbound HTTPS).             |
| Q5 pass                | The per-team boundary is real. The service-side allowlist is defence in depth.                                                |
| Q5 fail                | The app can read teams it is not installed in. Stop: the security model in the scope is wrong.                                |

## Clean-up

1. Delete the secret from the app registration and `unset APP_SECRET TOKEN BTOKEN`.
2. Uninstall the Teams app from `m365-bot-probe`, delete the team, and delete the test app and bot registrations.
3. Keep the filled-in results table; it becomes part of the scope document.
