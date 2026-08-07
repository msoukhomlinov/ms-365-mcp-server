# Agent gotchas: mail, calendar, and To Do tools

This page collects argument-shape mistakes an LLM agent commonly makes against this server's
mail, calendar, task, and contact tools, and the exact Microsoft Graph errors that result. It's
aimed at the model driving an MCP client, not at people modifying this codebase — point an
agent's system prompt or a project-level instructions file at this doc before its first call
into these tools.

Tool names below are this server's registered tool names (kebab-case), called with plain MCP
`{"name": "...", "arguments": {...}}` tool calls. Adjust for whatever prefix or namespacing your
own MCP client adds.

## Six mistakes that break things

1. **The id parameter is usually `messageId`, not `id` — same pattern for other resources.**
   Listings return a field called `id`; the tool that operates on one item wants that value
   under a resource-specific name: `messageId`, `mailFolderId`, `eventId`, `todoTaskListId`,
   `todoTaskId`, `contactId`, and so on. `list-mail-folder-messages` **requires**
   `mailFolderId` — there is no default. If a call is rejected as missing a required parameter,
   check you didn't send `id`.

2. **`expand` is for navigation properties only.** `body`, `sender`, `from` and `subject` are
   NOT expandable — ask for them in `select` instead. `expand` takes an array of strings;
   `select` takes a comma-separated string. Expanding a non-navigation property fails with Graph
   400 `Parsing OData Select and Expand failed: ... is not a navigation property or complex property. Only navigation properties can be expanded.`

3. **`search` takes ONE pair of double quotes around the WHOLE expression.**
   Right: `"search": "\"from:jane@example.com AND subject:flooring\""`
   Wrong: `"search": "\"from:jane@example.com\" AND subject:flooring"` → Graph 400
   `search` cannot be combined with `filter`, and `skip` does not work with `search`.

4. **`search` cannot be combined with `orderby` either — a separate error from the one above.**
   Right: drop `orderby` (results on messages are already sorted by sent date/time, newest
   first — not by relevance), or drop `search` in favor of `filter`. If you need the single most
   relevant match rather than the most recent one, narrow the KQL query instead of trying to sort
   for it — add more specific terms, `AND` extra clauses, or quote exact phrases.
   Wrong: passing both → Graph 400 `SearchWithOrderBy`: _"The query parameter '$orderBy' is not
   supported with '$search'."_

5. **With `filter` + `orderby` together (no `search`), the `orderby` property must come FIRST
   in `filter`.**
   Right: `filter: "receivedDateTime ge 2026-07-20T00:00:00Z"` with
   `orderby: "receivedDateTime desc"`
   Wrong: `filter: "from/emailAddress/address eq 'x@y.com' and receivedDateTime ge ..."` with
   `orderby: "receivedDateTime desc"` → Graph 400 `InefficientFilter`
   To find mail from one sender, prefer `search: "\"from:x@y.com\""` and drop `filter`/`orderby`
   entirely — or put `receivedDateTime` first in `filter`.

6. **`ErrorInvalidIdMalformed` almost always means the ID was mistyped, not a server bug.**
   `get-mail-message` and `list-mail-attachments` both work fine on a correctly-copied ID.
   Message and attachment IDs are long strings (100+ characters), often ending in `=` or `==`.
   Retyping one, or reusing one from earlier in the conversation instead of the listing result
   in front of you, can drop or add a character without it being obvious.
   **Rule: always copy the `id` field character-for-character from the listing result you just
   got. Never retype an ID. Never reuse an ID from an earlier turn.** On this error, don't retry
   the same ID — call `list-mail-messages` / `list-mail-folder-messages` again, get a fresh
   `id`, and use that exact value.

## Recipes

Recent inbox mail — `list-mail-messages` maps to `/me/messages` and searches across every
mailbox folder (sent items, drafts, everything), so it's the wrong tool when you actually want
the inbox. Use `list-mail-folder-messages` with `mailFolderId: "inbox"` instead:

```json
{
  "name": "list-mail-folder-messages",
  "arguments": {
    "mailFolderId": "inbox",
    "top": 20,
    "orderby": "receivedDateTime desc",
    "select": "id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments"
  }
}
```

Mail since a date — note `receivedDateTime` first in `filter`:

```json
{
  "name": "list-mail-messages",
  "arguments": {
    "top": 30,
    "filter": "receivedDateTime ge 2026-07-20T00:00:00Z",
    "orderby": "receivedDateTime desc",
    "select": "id,subject,from,receivedDateTime,bodyPreview,isRead"
  }
}
```

Find mail from a sender or about a topic — no `filter`, no `orderby`:

```json
{
  "name": "list-mail-folder-messages",
  "arguments": {
    "mailFolderId": "inbox",
    "search": "\"from:jane@example.com AND subject:flooring\"",
    "top": 10,
    "select": "id,subject,from,receivedDateTime,bodyPreview"
  }
}
```

Well-known `mailFolderId` values: `inbox`, `sentitems`, `drafts`, `deleteditems`, `archive`,
`junkemail`. For anything else, get real ids from `list-mail-folders`.

Read one message in full — pass the listing's `id` as `messageId`, and use `select` (never
`expand`) for the body:

```json
{
  "name": "get-mail-message",
  "arguments": {
    "messageId": "AAMkAGU3...",
    "select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments"
  }
}
```

Omitting `select` returns the whole message, which is also fine and often smaller than expected.

Reply, draft-first — get the user's approval before anything is sent. `comment` is not a
registered top-level parameter here; the tool's only declared argument is `body`, so a
top-level `comment` gets silently dropped and the draft is created with no reply text. Nest it
under `body` instead — same shape applies to `create-reply-all-draft` and `create-forward-draft`:

```json
{
  "name": "create-reply-draft",
  "arguments": {
    "messageId": "AAMkAGU3...",
    "body": { "comment": "..." },
    "confirm": true
  }
}
```

```json
{
  "name": "send-draft-message",
  "arguments": { "messageId": "<draft id>", "confirm": true }
}
```

Only call `send-draft-message` after the user has approved that specific draft.

Calendar for a window — `get-calendar-view` expands recurrences, `list-calendar-events` does not:

```json
{
  "name": "get-calendar-view",
  "arguments": {
    "startDateTime": "2026-07-27T00:00:00",
    "endDateTime": "2026-08-03T00:00:00",
    "timezone": "UTC",
    "top": 50,
    "select": "id,subject,start,end,location,organizer,attendees,isAllDay"
  }
}
```

To Do — always list the lists first, tasks live under a list:

```json
{ "name": "list-todo-task-lists", "arguments": {} }
```

```json
{
  "name": "list-todo-tasks",
  "arguments": { "todoTaskListId": "<id>", "top": 25, "filter": "status ne 'completed'" }
}
```

## Writes and the `confirm` gate

If the server is deployed with `MS365_MCP_REQUIRE_CONFIRM=true` (an opt-in, per-deployment
setting — off by default), every mail, calendar, contact, and task mutation is rejected with
`{"error": "confirmation_required"}` unless the call includes `confirm: true`. The parameter is
`confirm`, not `confirmed`. Get the user's approval for the specific action first; never set
`confirm: true` just to clear the error.

## Authentication

In stdio mode with local MSAL login, tokens are cached and normally persist across restarts.
Only if a call returns `Failed to acquire token` or `No accounts found. Please login first.`
(the latter is normal on a fresh install or right after logout, since the token cache starts
empty):

1. Call `verify-login` to confirm the actual state rather than assuming.
2. If the session really has expired, `login` returns a device code URL — the user has to open
   it and enter the code themselves; the server never opens a browser. This is the default,
   device-code flow.
3. Call `verify-login` again before retrying the original operation.

Don't call `login` pre-emptively before ordinary mail or calendar work.

If the server was started with `--auth-browser` instead, this is different: `login` calls
`acquireTokenInteractive()`, which opens the system browser itself for the user to sign in,
rather than returning a device-code URL. If you're talking to a server configured that way,
don't expect a device-code URL back from `login` — the browser flow handles the interaction
directly.

## Attachments — listing them, and getting at their content

`list-mail-attachments` works and is how you see what's on a message. **Always pass `select`:**

```json
{
  "name": "list-mail-attachments",
  "arguments": {
    "messageId": "AAMkAGU3...",
    "select": "id,name,contentType,size,isInline"
  }
}
```

One object per attachment: `id`, `name`, `contentType`, `size` (bytes, approximate), and
`isInline`. That's enough to know what arrived and which entries are inline signature images
rather than real files. **Read `size` before you try to move any bytes** — it decides which of
the options below is worth attempting at all.

**This result has no `messageId` field — it's attachment objects only (`id`, `name`,
`contentType`, `size`, `isInline`).** Keep using the same `messageId` you already passed in to
make this call; don't retype it. Copy the attachment `id` exactly, character-for-character, from
this result — don't retype or shorten it, and don't reuse one from earlier in the conversation,
get a fresh one from the listing you just made. See mistake 6.

**This server does not convert documents to text.** No tool here turns a PDF, Word, PowerPoint,
Excel, OpenDocument or RTF attachment into markdown, so don't tell the user you can read one.
Reaching an attachment's content means moving its raw bytes somewhere, and which of the three
tools below can do that depends on how the server was started.

### `get-download-url` — a URL fetched out-of-band

Always registered, but what it can actually resolve varies by resource and by deployment:

- **OneDrive and SharePoint files: always works, no configuration.** It returns Graph's own
  pre-authenticated `@microsoft.graph.downloadUrl` for a driveItem, which streams the bytes with
  no `Authorization` header. Pass a drive item path — `/me/drive/items/{driveItem-id}`,
  `/drives/{drive-id}/items/{driveItem-id}/content`, or
  `/me/drive/root:/path/file.pdf:/content`. A trailing `/content` is optional and is stripped
  for you. Prefer this over `download-bytes` for any drive file above a few KB.
- **Mail attachments, event attachments, meeting recordings and other `/$value` endpoints: only
  if the operator opted in.** Graph publishes no pre-authenticated URL for these, so the server
  has to mint one of its own — and it only does that when started with
  `--enable-attachment-urls`, **in HTTP mode**, using credentials the server itself holds.

Where minting is enabled, the same `/$value` path you'd hand `download-bytes` returns a URL
served by this server rather than by Graph:

```json
{
  "name": "get-download-url",
  "arguments": {
    "target": "/me/messages/AAMkAGU3.../attachments/AAMkAGU3.../$value"
  }
}
```

→ `{ downloadUrl, expiresAt, singleUse: true, note }`. It is good for exactly **one** fetch and
expires quickly (120 seconds by default, 300 at most), so pass it straight to whatever will
fetch it. Don't sit on it, don't log it for later, and don't retry a URL some fetch has already
spent — mint a fresh one.

Where minting is off, you get an error naming the resource kind instead of a URL:

```text
Mail and calendar event attachments do not expose a pre-authenticated download URL. Use download-bytes for small attachments.

Meeting recordings do not expose a pre-authenticated download URL. Use download-bytes for small recordings or get-meeting-recording-content where available.

$value byte endpoints do not expose a pre-authenticated download URL. Use download-bytes to read these bytes.
```

And on a server that _does_ pass the flag but takes its Graph identity from the request — plain
HTTP bearer mode, `--obo`, or OAuth — minting is refused on purpose:

```text
Server-minted download URLs are unavailable when Graph identity comes from the request (OAuth, OBO, or bearer mode): the URL is redeemed later without an Authorization header, so the bytes would be fetched as a different identity than the one that asked for them. Use download-bytes.
```

Read any of those four as "this deployment can't give me a URL for this resource," and move on to
the options below. Don't re-send the same call.

**Don't rule this tool out from its own description.** The `get-download-url` description states
flatly that mail attachments and meeting recordings do not expose a pre-authenticated URL and
points you at `download-bytes`; it describes Graph's behaviour and doesn't mention the minting
flag. On a server started with `--enable-attachment-urls` the call succeeds regardless. One
attempt is cheap and the refusal above is short and unambiguous.

### `download-bytes-to-file` — straight to the server's disk (stdio only)

For a file the user wants **saved** rather than read, this writes the authenticated bytes to an
absolute `outputPath` on the server's own filesystem and returns
`{ path, contentType, bytesWritten }` — never base64 through your context, so it's the right
choice regardless of file size. It won't overwrite an existing file. **stdio deployments only:
it isn't registered at all over HTTP.**

That makes it the mirror image of minting, and the pairing is worth remembering: minted
attachment URLs are HTTP-mode-only, `download-bytes-to-file` is stdio-only. No single deployment
offers both, so for a mail attachment's bytes at most one out-of-band route exists on whatever
server you're talking to.

### `download-bytes` — base64 into your context, last resort

Returns `{ contentType, encoding: "base64", contentLength, contentBytes }` — the whole file,
inline, in your own context. Base64 inflates it by about a third, and the result is only useful
to you if the underlying bytes are already text.

Call it only when the file is **both small and genuinely plain text** (`.txt`, `.csv`, `.md`,
`.json`, `.xml`, `.log`), or when the user has explicitly asked for raw bytes knowing what they
are. A few hundred KB of PDF costs tens of thousands of tokens and yields nothing readable.

**One trap:** this tool's own description tells you that for large files you should "prefer
`get-download-url`". For a drive or SharePoint file that's correct — follow it. For a mail
attachment or a recording on a deployment that didn't pass `--enable-attachment-urls`, it's a
dead end: `get-download-url` refuses with one of the errors above and you end up back here.

- **Avoid `get-mail-message-mime` for reaching an attachment.** It returns the whole RFC 5322
  message with every attachment base64-inline, so a modest attachment can balloon the response
  well past its own file size once base64-encoded and wrapped in the surrounding MIME structure.
- **Avoid dropping `select` on a message just to get at an attachment's content.** Without
  `select`, `list-mail-attachments` includes `contentBytes` — the entire file as base64, inline
  in the listing result. That's the same cost as `download-bytes` with no chance to read `size`
  and back out first. Use `get-download-url` (out-of-band, where the deployment supports it),
  `download-bytes-to-file` (saving to disk, stdio only), or `download-bytes` (small plain-text
  files read into context) on the attachment's own `/$value` path instead; all three take the
  attachment `id` from a normal, `select`ed listing call.

### What to say when you can't read an attachment

For document formats this is the **common** case, not the exception: nothing in this server
extracts text, so a PDF, Word, PowerPoint or Excel attachment is unreadable to you unless
something outside the server converts it. Say so in one sentence, name the file and its size,
and offer the alternative:

> There's one attachment — `Quote-1042.pdf`, 191 KB. I can't read PDF content — this server hands
> me raw bytes, not text. Paste the part you care about and I'll work from that.

Never pull a document down with `download-bytes` and then guess at its contents from the base64.
If you can't read it, say so.

## When a call fails

- `-32602 Input validation error` → the arguments didn't validate against the tool's schema at
  all — the SDK returns this same code whether the problem is a wrong parameter **name** (see
  mistake 1), a wrong **type** for a correctly-named parameter (e.g. passing `expand` as a plain
  string instead of an array of strings — see mistake 2), a missing required field, or a value
  outside an allowed set. Don't assume it's always the mistake-1 naming issue and reflexively
  rename something. Read the validation details in the error response, or fetch the tool's
  schema once, to see the actual names, required fields, types, and allowed values. Don't
  re-send the same arguments unchanged.
- Graph 400 `Parsing OData Select and Expand failed` / `Only navigation properties can be expanded` → mistake 2.
- Graph 400 with a `$search` quoting complaint → mistake 3.
- Graph 400 `SearchWithOrderBy` → mistake 4.
- Graph 400 `InefficientFilter` → mistake 5.
- Graph 400 `ErrorInvalidIdMalformed` → the ID was mistyped. See mistake 6. Get a fresh `id`
  from `list-mail-messages` / `list-mail-folder-messages` and copy it exactly — don't retry the
  same ID, don't reuse one from earlier in the conversation.
- Graph 400 `Could not find a property named '...' on type 'microsoft.graph.attachment'` → a
  non-selectable field was passed in `select`; see the attachments section above.
