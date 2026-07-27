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
   400 `Parsing OData Select and Expand failed: ... is not a navigation property or complex
   property. Only navigation properties can be expanded.`

3. **`search` takes ONE pair of double quotes around the WHOLE expression.**
   Right: `"search": "\"from:jane@example.com AND subject:flooring\""`
   Wrong: `"search": "\"from:jane@example.com\" AND subject:flooring"` → Graph 400
   `search` cannot be combined with `filter`, and `skip` does not work with `search`.

4. **`search` cannot be combined with `orderby` either — a separate error from the one above.**
   Right: drop `orderby` and rely on relevance ranking, or drop `search` in favor of `filter`.
   Wrong: passing both → Graph 400 `SearchWithOrderBy`: *"The query parameter '$orderBy' is not
   supported with '$search'."*

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

Reply, draft-first — get the user's approval before anything is sent:

```json
{
  "name": "create-reply-draft",
  "arguments": { "messageId": "AAMkAGU3...", "comment": "...", "confirm": true }
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
Only if a call returns `Failed to acquire token`:

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

## Attachments — read documents directly with `convert-document`

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
rather than real files.

**Copy the `messageId` and attachment `id` exactly, character-for-character, from this result.**
Don't retype or shorten them, and don't reuse an ID from earlier in the conversation — get a
fresh one from the listing you just made. See mistake 6.

**To read a PDF, Word, PowerPoint, Excel, OpenDocument, or RTF attachment, call
`convert-document` on its `/$value` path — it returns markdown text, not bytes you'd have to
decode and parse yourself:**

```json
{
  "name": "convert-document",
  "arguments": {
    "target": "/me/messages/AAMkAGU3.../attachments/AAMkAGU3.../$value"
  }
}
```

`convert-document` only exists if the server was started with `--enable-document-conversion`
(off by default — see the README's "Document Conversion" section for the rest of the
prerequisites). If it isn't registered, fall back to `download-bytes` and note that the caller
will get raw base64 to decode itself, not text.

Where it is available, `convert-document` is dramatically more token-efficient than pulling the
same file as inline base64 through `list-mail-attachments` — a multi-hundred-KB document becomes
a few thousand tokens of markdown instead of tens of thousands of tokens of unreadable base64.
Source size is capped at 25 MB, checked against the real decoded bytes rather than a reported
size; over that, `convert-document` returns a clear error instead of buffering the whole file.
`truncated` / `totalLength` in the result tell you if the markdown itself was cut down further
(default cap 20,000 characters) — ask for OCR (`ocr: true`) only for scanned/image-based
documents; it's off by default because it adds real latency.

For a format `convert-document` doesn't handle (images, unusual binary types), or a file you
want saved rather than read, `download-bytes` on the same `/$value` path returns raw bytes as
base64 — usable, but only worth calling for genuinely small files, since the bytes land in your
own context as base64. For document *content*, `convert-document` (where enabled) is almost
always the right tool instead.

- **Avoid `get-mail-message-mime` for reaching an attachment.** It returns the whole RFC 5322
  message with every attachment base64-inline, so a modest attachment can balloon the response
  well past its own file size once base64-encoded and wrapped in the surrounding MIME structure.
- **Avoid dropping `select` on a message just to get at an attachment's content.** Without
  `select`, `list-mail-attachments` includes `contentBytes` — the entire file as base64, inline
  in the listing result. Use `convert-document` (documents) or `download-bytes` (small
  non-document files) with the attachment's own `/$value` path instead; both take the
  attachment `id` from a normal, `select`ed listing call.

### What to say when a format truly can't be read

This only applies to formats `convert-document` doesn't cover (images, most non-document binary
types), a source over the 25 MB cap, or a deployment where document conversion isn't enabled at
all. One sentence, then offer the alternative:

> There's one attachment — `photo.jpg`, 4 MB. I can't extract text from an image attachment.
> Describe what you need from it and I'll work from that.

## When a call fails

- `-32602 Input validation error` → wrong parameter **name**. See mistake 1, or fetch the tool's
  schema once. Don't re-send the same arguments unchanged.
- Graph 400 `Parsing OData Select and Expand failed` / `Only navigation properties can be
  expanded` → mistake 2.
- Graph 400 with a `$search` quoting complaint → mistake 3.
- Graph 400 `SearchWithOrderBy` → mistake 4.
- Graph 400 `InefficientFilter` → mistake 5.
- Graph 400 `ErrorInvalidIdMalformed` → the ID was mistyped. See mistake 6. Get a fresh `id`
  from `list-mail-messages` / `list-mail-folder-messages` and copy it exactly — don't retry the
  same ID, don't reuse one from earlier in the conversation.
- Graph 400 `Could not find a property named '...' on type 'microsoft.graph.attachment'` → a
  non-selectable field was passed in `select`; see the attachments section above.
