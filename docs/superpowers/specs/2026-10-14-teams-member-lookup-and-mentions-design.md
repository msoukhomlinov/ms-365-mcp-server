# Teams member lookup and mention shorthand

Two independent upstream contributions to Softeria `ms-365-mcp-server`. Each branches from `upstream/main`, ships as its own PR, and does not depend on the other.

## Background

A comparison with `InditexTech/mcp-teams-server` showed two Teams ergonomics gaps in this server:

1. Finding a person by name means listing every member and filtering in the model.
2. Mentioning someone means hand-building numbered `<at id="N">` tags plus a matching `mentions` array. Mismatches return Graph 400s.

This spec covers those two gaps only. Out of scope:

- Name-to-id resolution. This is a possible follow-up built on A plus B.
- Channel, team and tag mentions.
- Default team/channel pinning. Dropped.
- Scoped app-only (client credentials, RSC) auth. This is a separate project, brainstormed on its own.

## Success criteria

- A maintainer can review and merge either PR without the other.
- No existing tool name, required parameter or correct request body changes behavior.
- Both changes follow upstream patterns: `endpoints.json` entries and `llmTip`, request rewriting in `executeGraphTool` in `src/graph-tools.ts`, tests in `src/__tests__/graph-tools.test.ts`.

## PR A: `name` filter on member lists

Tools: `list-team-members`, `list-chat-members`.

- New optional `name` string parameter on those two tools only. It is removed before the Graph request is built, as `confirm` is, and never sent to Graph.
- Matching: case-insensitive substring match of `name` against `displayName` and `email` of each item in the response `value` array.
- The filter runs on the final response body, after `$select` projection and the user-field restrictions. It matches only on fields the caller can see. Matching on a hidden field would let the filter act as an oracle for that field.
- Pagination: with `fetchAllPages` the filter runs after pages are merged. Without it, only the fetched page is filtered, and the parameter description says so.
- No match returns an empty `value` array, not an error.
- `llmTip` on both endpoints documents `name` and recommends it before mentioning someone.

## PR B: mention shorthand

Tools (seven; create, reply and update for channel and chat): `send-channel-message`, `reply-to-channel-message`, `update-channel-message`, `update-channel-message-reply`, `send-chat-message`, `reply-to-chat-message`, `update-chat-message`.

The Planner task-chat post tool also accepts mentions, but with a different body shape (`mentioned: 'user-id'`). It is out of scope.

- The model writes `<at user-id="<aad-guid>">Jane Doe</at>` inside html content.
- The server rewrites each tag to `<at id="N">Jane Doe</at>`. N starts above the highest numeric `<at id>` already in the content, so hand-built tags never collide.
- For each rewritten tag it appends `{ id: N, mentionText: "Jane Doe", mentioned: { user: { id: "<guid>", displayName: "Jane Doe", userIdentityType: "aadUser" } } }` to `mentions`. If the caller already supplied a `mentions` array, new entries are merged into it.
- A body that already uses numeric `<at id>` tags and its own `mentions` array passes through untouched.
- A `user-id` that is not a GUID returns a tool error naming the bad value. A literal tag is never sent to Graph.
- Only `contentType: "html"` is rewritten. Other content types are left alone.
- `llmTip` on each affected tool shows the shorthand first and keeps the raw form as the fallback.

## Tests

Written first, in `src/__tests__/graph-tools.test.ts`, mirroring the existing `$search` normalization tests.

PR A:
- match on name
- match on email
- no match returns empty `value`
- `name` is not forwarded to Graph
- filter runs after the `fetchAllPages` merge
- filter runs after projection, and a field removed by projection is not matched

PR B:
- single mention
- multiple mentions
- id collision with hand-built mentions
- correct hand-built body passes through unchanged
- bad GUID returns a tool error
- non-html content untouched
- rewrite applies on each affected tool

## Process

- Branch each PR from `upstream/main`. The spec file stays on the fork integration branch and is not part of either PR.
- Each PR runs the full test suite before it is opened.
- No AI attribution in commits, PR text or comments.
