# E2E runbook: Outlook attachment to Markdown

Proves: M365 mail attachment -> ms-365-mcp-server (`--attachment-proxy`) -> ticket URL -> docglean-mcp -> Markdown.

The default (unsigned, Docker) path — sections 0 through 5, plus the slicing/error checks in 5.3/6 and the network-isolation checks in 7.1 — was run end to end on 2026-10-11 against a real mailbox, using generated (not hand-crafted-in-Outlook) test attachments; results are folded into the sections below. Steps still marked "not verified" are paths that run was not required to exercise: the plain-`node` path (2b), the keyed variant (Appendix A), and the ticket-reuse/expiry checks in 7.2 that need stopping and swapping the compose profile. Never paste keys, tokens or device codes into chat or tickets.

It follows the `ms-365-mcp-server` source on this branch. The default (unsigned) setup needs a `docglean-mcp` build that includes the unsigned-ticket change, which is not yet released; the keyed appendix works against a docglean `main` build. For the configuration behind it, see "Reading documents through a conversion proxy" in the [README](../README.md#reading-documents-through-a-conversion-proxy).

## 0. Prerequisites

- Docker with Compose v2 (docglean is Linux-only; it runs only as a container).
- Node 22+ only for the plain-`node` path (section 2b).
- An M365 mailbox you can send mail to yourself from. Delegated `Mail.Read` is requested at login. The built-in app registration is used unless you set `MS365_MCP_CLIENT_ID`; a work tenant may need admin consent. A personal (MSA) account needs `MS365_MCP_TENANT_ID=consumers`.
- The server's own login. `read-document` refuses OAuth, OBO and bearer identity (`identity_not_supported`), so the run uses the server's cached device-code account with `--trust-proxy-auth`. You log in once in step 2.
- Both repos checked out: this one on a branch that has `--attachment-proxy`, and `docglean-mcp`. The scanned-PDF check depends on whether your docglean build has OCR (section 6).

Work directory: the compose file is [`e2e/docker-compose.yml`](e2e/docker-compose.yml). `.env` and `secrets/` are created beside it and are ignored by version control there. From the root of this repo:

```bash
cd docs/e2e
```

## 1. Tokens (files only), and which signing mode

Two modes pair ms-365 with a docglean profile. They must match; a mismatch is `blocked_host` on every read.

| ms-365 setting                                | URL it mints                      | docglean profile that accepts it                                                     |
| --------------------------------------------- | --------------------------------- | ------------------------------------------------------------------------------------ |
| No `MS365_MCP_ATTACHMENT_URL_KEY`/`_KEY_FILE` | `/attachment?t=<43-char ticket>`  | Unsigned: `_ORIGIN`, `_ALLOW_CIDRS` (one `/32`), `_UNSIGNED_TICKET_PATH=/attachment` |
| Key set                                       | the same plus `dgk`, `dgx`, `dgs` | Keyed: `_ORIGIN`, `_ALLOW_CIDRS`, `_KEY_FILE`, `_KEY_ID`                             |

- A keyed ms-365 against a keyless docglean profile is rejected: that profile refuses any URL carrying `dgk`, `dgx` or `dgs`.
- A keyless ms-365 against a keyed-only docglean profile is `blocked_host`: the bare URL is not signed, so it is never dialled. A docglean profile with both a key and the unsigned path accepts both (a migration window).

This walkthrough uses the unsigned mode, so there is no key to generate. The unsigned profile needs a docglean build that includes the unsigned-ticket change (not yet released). For the keyed mode see Appendix A.

```bash
umask 077
cat > .env <<EOF
DOCGLEAN_REPO=/path/to/docglean-mcp
DOCGLEAN_TOKEN=$(openssl rand -hex 32)
EOF
# personal Microsoft account only:  echo MS365_MCP_TENANT_ID=consumers >> .env
```

Notes:

- `DOCGLEAN_TOKEN` is the bearer docglean demands because it binds `0.0.0.0` (it refuses to start with none). ms-365 sends it via `MS365_MCP_ATTACHMENT_PROXY_TOKEN`. Neither side offers a `_FILE` form for this one, hence `.env`. Keep `.env` out of git.
- Throwaway for this test; delete it at teardown.
- Without a signature the ticket is the only credential, so it is single-use (3 fetches for `read-document`) and short-lived, and the docglean `_ALLOW_CIDRS` pin (a single address, the attachment listener) is what keeps the fetch on that listener.
- The unsigned profile requires the docglean origin to equal the URL base (`http://m365:3001`) and the path to be exactly `/attachment`. No other query parameter is accepted.

## 2a. Bring up (Docker path)

```bash
docker compose build                      # builds m365-mcp:e2e and docglean-mcp:e2e locally; nothing is pushed
                                          # (ms-365 build runs `npm run generate`, which needs network; confirmed working)
docker compose up -d docglean             # healthcheck start period allows up to ~90 s; observed healthy in ~7 s
```

Log in the server's own account once (the cache lands in the `m365-data` volume, file-backed, keytar off):

```bash
docker compose run --rm --no-deps m365 --login --preset mail
# prints a URL and a code; open the URL in a browser and enter the code. Nothing to paste back.
docker compose run --rm --no-deps m365 --list-accounts
docker compose run --rm --no-deps m365 --verify-login --preset mail
```

Then start the server:

```bash
docker compose up -d m365
docker compose logs m365 | grep -E "listening|MCP endpoint|Document proxy|Attachment"
```

Expected startup lines (wording may differ slightly):

- `Server listening on 10.89.0.2:3000` and `MCP endpoint: http://10.89.0.2:3000/mcp`
- `Document proxy: http://docglean:8080/mcp (read-document only; ...bearer credential set)`
- `Attachment URLs: http://m365:3001/attachment (ttl 120s, unsigned)`, or `(ttl 120s, signed, key id 1)` in the keyed variant
- `Attachment listener: separate, port 3001 ...`
- No warning about both listeners sharing an interface. If you see one, `--attachment-host` is not taking effect.

None of these lines appeared in `docker compose logs m365` at the default log level in the 2026-10-11 run — the container was otherwise healthy and correct (confirmed via the health checks below and a full, successful run), so treat their absence as a log-verbosity gap, not a startup failure. Use the health checks in section 3 to confirm readiness instead of grepping for these lines.

Startup refusals and what they mean:

- `MS365_MCP_ATTACHMENT_URL_KEY_FILE could not be read` / `is empty`: only if you set a key file; unset it for the unsigned mode.
- `MS365_MCP_ATTACHMENT_URL_BASE` required / invalid / IPv6 literal (only refused when a key is set).
- `--attachment-port requires --enable-attachment-urls` / `--attachment-host requires --attachment-port`.
- Proxy flag with `--toon` is refused.
- Stdio mode ignores the flag with a warning; the proxy is HTTP only.

## 2b. Plain `node` path (smoke test of the code path only)

Docglean still runs in Docker. Because ms-365 now runs on the host, the network isolation is NOT demonstrated here; use 2a for that.

```bash
cd ../..                                         # the root of this repo
npm ci && npm run generate && npm run build
node dist/index.js --login --preset mail         # device code; macOS Keychain or file cache
```

Pick how the container reaches the host, then make docglean match it (`host.docker.internal` is the Docker Desktop name; on Linux add `--add-host host.docker.internal:host-gateway`):

```bash
cd docs/e2e
H=$(docker run --rm docglean-mcp:e2e python -c "import socket;print(socket.gethostbyname('host.docker.internal'))")
docker run -d --name docglean-e2e -p 127.0.0.1:8080:8080 \
  --add-host host.docker.internal:host-gateway \
  --env-file <(grep DOCGLEAN_TOKEN .env | sed 's/^DOCGLEAN_TOKEN/DOCGLEAN_MCP_BEARER_TOKEN/') \
  -e DOCGLEAN_MCP_SIGNED_ORIGIN_M365_ORIGIN=http://host.docker.internal:3001 \
  -e DOCGLEAN_MCP_SIGNED_ORIGIN_M365_ALLOW_CIDRS="[\"$H/32\"]" \
  -e DOCGLEAN_MCP_SIGNED_ORIGIN_M365_UNSIGNED_TICKET_PATH=/attachment \
  -e DOCGLEAN_MCP_ALLOW_INSECURE_SIGNED_ORIGIN=true docglean-mcp:e2e

export MS365_MCP_ATTACHMENT_URL_BASE=http://host.docker.internal:3001
export MS365_MCP_ATTACHMENT_PROXY_TOKEN=$(grep ^DOCGLEAN_TOKEN .env | cut -d= -f2)
cd ../..
node dist/index.js --http 127.0.0.1:3000 --trust-proxy-auth --preset mail \
  --attachment-proxy http://127.0.0.1:8080/mcp --attachment-port 3001 --attachment-host 0.0.0.0
```

Caveats **(not verified)**: the host address Docker Desktop presents for `host.docker.internal` may differ from what a container sees as the peer; the `/mcp` loopback plus wildcard attachment bind triggers the server's shared-interface warning; the attachment port is then reachable from your LAN. Use `http://127.0.0.1:3000/mcp` for clients in this path. Skip steps 3 and 7's network commands; they assume 2a.

## 3. Health checks

```bash
docker compose ps                                   # docglean: healthy; m365: healthy
# docglean (open paths, no bearer):  /healthz = liveness, /readyz = 200 once workers are warm
docker compose exec -T docglean python -c "import urllib.request as u;print(u.urlopen('http://127.0.0.1:8080/readyz').status)"
# m365 MCP port health route (returns 'Microsoft 365 MCP Server is running'):
docker compose exec -T m365 node -e "fetch('http://10.89.0.2:3000/').then(async r=>console.log(r.status,await r.text()))"
```

Helper for every MCP call below (the m365 container can reach its own agent-net address; no host port is needed):

```bash
mcp() { # usage: mcp '<json-rpc body>'
  docker compose exec -T -e BODY="$1" m365 node -e '
    fetch("http://10.89.0.2:3000/mcp",{method:"POST",
      headers:{"content-type":"application/json",accept:"application/json, text/event-stream"},
      body:process.env.BODY}).then(async r=>{console.log("HTTP",r.status);console.log(await r.text())})'
}
mcp '{"jsonrpc":"2.0","id":1,"method":"tools/list"}' | head -c 3000
```

Expected tool list (preset `mail`, proxy on): `read-document` present; `download-bytes`, `get-download-url`, `get-mail-message-mime` absent — confirmed exactly on 2026-10-11 (44 tools registered, including `send-mail`, `get-upload-url`). The server is stateless (no session); `tools/list` and `tools/call` were both called directly with no `initialize` and worked.

Optional real client from your Mac instead of the helper: uncomment `ports:` on `m365` in the compose file and use `claude mcp add --transport http ms365-e2e http://127.0.0.1:3000/mcp` or `npx @modelcontextprotocol/inspector` (Streamable HTTP, URL `http://127.0.0.1:3000/mcp`). Whether Docker forwards a published port to a two-network container bound to one address is **not verified**; if it times out, keep the helper. No Authorization header is needed under `--trust-proxy-auth`.

## 4. Create the test mails (send to yourself, wait until they are in the Inbox)

1. **pdf-text**: subject `E2E pdf-text`. Attach a PDF with a real text layer, a few pages (export from Word/Pages "Save as PDF"; confirm you can select text in it). Put a unique phrase on page 1 and another on the last page.
2. **docx**: subject `E2E docx`. Attach a `.docx` with a heading, a bullet list and a small table.
3. **item attachment (rfc822)**: subject `E2E item`. In Outlook, open any received message, choose Forward as attachment (OWA/new Outlook: `...` menu, Forward as attachment; classic: Home, More, Forward, As Attachment, or Insert, Attach Item, Outlook Item on a new message). Send it to yourself. Graph models that as an `itemAttachment`, not a file. For the contrast case also send a message with a `.eml` file saved to disk attached (a `fileAttachment` with type `message/rfc822`).
4. **scanned**: subject `E2E scan`. Attach an image-only PDF: print a page to PNG or screenshot it, then export that image as PDF so there is no text layer (select-all should select nothing). Do not OCR it first.

Verified 2026-10-11, programmatically rather than through the Outlook UI: a real 4-page PDF built with a genuine `Tj`-operator text layer (`cupsfilter` on plain text with form-feed page breaks), a real `.docx` built as raw OOXML (heading style, `numPr` bullet list, `w:tbl` table — no Word/python-docx involved), an image-only PDF with no `/Font` dictionary (Pillow-drawn PNG wrapped via `sips`), and a hand-built RFC 5322 `.eml` for the contrast case. For the item-attachment case, rather than driving Outlook's UI, a `send-mail` call was made directly with an attachment of `"@odata.type": "#microsoft.graph.itemAttachment"` carrying an embedded `#microsoft.graph.message` object — Graph accepted it and `list-mail-attachments` reported it back with `contentType: "message/rfc822"`, the same shape a real "Forward as attachment" produces. All four — and the `.eml` contrast case — converted correctly; see section 6.

## 5. MCP calls

Argument names are camelCase/bare (`messageId`, `top`, `select`); `tools/list` is authoritative. Kebab forms like `message-id` are also accepted by the server.

### 5.1 Find the messages

```bash
mcp '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"list-mail-messages","arguments":{"top":10,"select":"id,subject,hasAttachments,receivedDateTime","search":"\"subject:E2E\""}}}'
```

If `search` is rejected or noisy, drop it and use `top` only. Copy each message `id` (long base64-ish string; keep it exact). In JSON output, `\/` is not used; copy as shown.

### 5.2 List attachments

```bash
mcp '{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"list-mail-attachments","arguments":{"messageId":"<MESSAGE_ID>","select":"id,name,contentType,size,isInline"}}}'
```

`@odata.type` tells you `#microsoft.graph.fileAttachment` vs `#microsoft.graph.itemAttachment`. Copy `<ATTACHMENT_ID>`.

### 5.3 Read it

`target` is a relative Graph path ending in `/$value` (bare attachment path is the JSON metadata resource and is refused with `invalid_target`). Do not URL-encode the ids unless they already contain `%`; percent-encoded separators (`%2F`) are refused.

```bash
mcp '{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"read-document","arguments":{"target":"/me/messages/<MESSAGE_ID>/attachments/<ATTACHMENT_ID>/$value"}}}'
```

Slicing (passed straight to docglean):

```json
{"target":"/me/messages/<MESSAGE_ID>/attachments/<ATTACHMENT_ID>/$value","pages":"1-2"}
{"target":"...","offset":20000,"maxChars":5000}
```

Whole message as RFC 5322 source: `"target":"/me/messages/<MESSAGE_ID>/$value"`.

Also valid: `/drives/{drive-id}/items/{item-id}/content` (OneDrive/SharePoint). The `mail` preset does not include drive listing tools, so use a separate preset for that; out of scope here.

## 6. Expected results

Success: the tool result is plain Markdown text (not JSON, no base64). docglean caps output at 20,000 characters per call (`max_chars` default), input at 25 MB (`DOCGLEAN_MCP_MAX_INPUT_BYTES`, default 25 MB), 500 PDF pages.

| Mail                   | Expect                                                                                                                                                                                                                                                                                                                            |
| ---------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| pdf-text               | Both unique phrases present — **confirmed**. `pages:"1"` returns only page 1 — **confirmed**. `pages` beyond the end (document has 4 pages; tried `pages:"99"`) returns `proxy_error` with `invalid_pages` in the message ("That selection lies entirely outside the document, which has 4 page(s)...") — **not** `no_text_in_selection` as this row previously claimed; that code is for `offset`/`maxChars` selecting nothing (not independently verified here). |
| docx                   | Heading as `#`, bullets as `-`, table as pipe table — **confirmed** exactly. `pages` is rejected (docx has no page unit): use `offset`/`maxChars`. Confirmed `proxy_error` with `pages_unsupported` in the message, word-for-word: "docx has no page, slide or sheet unit. Use offset to read it in slices."                     |
| item attachment        | Markdown of the forwarded message (headers, body) — **confirmed**, via a programmatically-built itemAttachment rather than Outlook's UI (see section 4); output was a `#`-heading, a bulleted Subject/From/To block, then the body text. Not confirmed: an item attachment that itself carries sub-attachments.                  |
| `.eml` file attachment | Same converter path as above — **confirmed**: same heading/metadata/body shape, docglean's `message/rfc822` support works.                                                                                                                                                                                                       |
| scanned PDF            | Without OCR: error `no_text_content` with `contentStatus: "no_text_layer"` (not exercised in this run — the docglean build used has OCR on by default). With OCR: **confirmed** — the image-only test PDF (no `/Font` dict, confirmed before sending) came back as clean Markdown with the embedded marker text recovered essentially perfectly, not merely "possibly imperfect". |

Paging caveat: `read-document` returns only the Markdown. docglean's `next_offset`/total-size fields are not shown to you by this tool, so continue by hand: `offset` = previous `offset` + characters you received. Every continuation mints a fresh ticket and re-fetches and re-converts the whole document.

### Failure modes

Error results are JSON text: `{"error":"<code>","message":...,"name":...,"contentType":...,"size":...}` with `isError:true`. The attachment facts (`name`, `contentType`, `size`) come from Graph and are null when unknown.

Only these docglean codes pass through under their own name: `unsupported_format`, `too_large`, `password_required`, `conversion_failed`, `fetch_failed`. Every other docglean code (including `blocked_host`, `timeout`, `busy`, `invalid_pages`, `pages_unsupported`) arrives as `error:"proxy_error"` with the raw code at the start of `message` (`blocked_host: ...`) and in `proxyCode`.

| Symptom                                                                | Cause                                                                                                                                                                                                                                                                                                                                                                                                                 | Fix                                                                                                                                                                                                                                                                                  |
| ---------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `proxy_error`, `blocked_host`                                          | docglean did not accept the URL or the address: signing-mode mismatch (keyed ms-365 against a keyless docglean profile, or keyless ms-365 against a keyed-only profile), origin not equal to URL base, wrong/missing key file, key id mismatch, expiry out of window (clock skew over 5 s, TTL over docglean max), or resolved IP not in `_ALLOW_CIDRS`. docglean deliberately gives the same answer for all of them. | First check both sides use the same mode (section 1). Compare `MS365_MCP_ATTACHMENT_URL_BASE` with `..._ORIGIN` character by character; confirm `docker compose exec docglean getent hosts m365` is `10.89.1.2`; same key file in both (keyed mode); `docker compose logs docglean`. |
| `fetch_failed` on a first read                                         | docglean reached m365 but m365 answered non-200: ticket spent/expired (404 is uniform), or Graph error (502)                                                                                                                                                                                                                                                                                                          | Retry the call (each call mints a fresh ticket). Check `docker compose logs m365`. A 502 on redemption usually means the cached login expired: redo `--login`.                                                                                                                       |
| `fetch_failed` on a repeat of the same URL                             | ticket spent: `read-document` tickets allow 3 fetches, counting failures, one at a time                                                                                                                                                                                                                                                                                                                               | Mint a new one by calling `read-document` again. Only reproducible directly: see 7.                                                                                                                                                                                                  |
| `proxy_unreachable`                                                    | m365 cannot reach `http://docglean:8080/mcp` (down, still warming, wrong name). The server retries once with a fresh mint                                                                                                                                                                                                                                                                                             | `docker compose ps`; `/readyz`; check `MS365_MCP_ATTACHMENT_PROXY` URL ends in `/mcp`. Wrong bearer gives `proxy_error` ("proxy refused the call"), not this.                                                                                                                        |
| `identity_not_supported`                                               | server not running on its own cached account: OAuth/OBO/bearer mode, or a bearer header reached the tool                                                                                                                                                                                                                                                                                                              | Use `--trust-proxy-auth`, no OBO, no `MS365_MCP_OAUTH_TOKEN`.                                                                                                                                                                                                                        |
| `invalid_target`                                                       | target not a plain `/$value` or `/content` Graph path (absolute URL, query, fragment, `..`, encoded slash), or mint failed                                                                                                                                                                                                                                                                                            | Use the shape in 5.3. Result message echoes the target.                                                                                                                                                                                                                              |
| `no_capacity`                                                          | 256 live tickets held                                                                                                                                                                                                                                                                                                                                                                                                 | Wait for TTL.                                                                                                                                                                                                                                                                        |
| `unsupported_format`                                                   | docglean could not identify the type (it uses the URL extension, then Content-Disposition filename, then Content-Type; ms-365 forwards Graph's type and filename)                                                                                                                                                                                                                                                     | Check `contentType`/`name` in the error; unsupported types such as images are expected here.                                                                                                                                                                                         |
| `too_large`                                                            | over 25 MB, or converted output limits                                                                                                                                                                                                                                                                                                                                                                                | Smaller file; or raise docglean's cap deliberately.                                                                                                                                                                                                                                  |
| `password_required`                                                    | encrypted file                                                                                                                                                                                                                                                                                                                                                                                                        | Expected for protected docs.                                                                                                                                                                                                                                                         |
| `no_text_content` (`contentStatus` `no_text_layer` / `empty_document`) | conversion worked, no text                                                                                                                                                                                                                                                                                                                                                                                            | Scan: expected before OCR. Blank: expected.                                                                                                                                                                                                                                          |
| `no_text_in_selection`                                                 | `pages`/`offset` selected nothing                                                                                                                                                                                                                                                                                                                                                                                     | You paged past the end; stop.                                                                                                                                                                                                                                                        |
| `expand_not_allowed`                                                   | a `$expand` of attachments was attempted on a list call                                                                                                                                                                                                                                                                                                                                                               | Do not expand; use `list-mail-attachments`.                                                                                                                                                                                                                                          |

## 7. Negative checks

### 7.1 Converter cannot reach `/mcp`, only the attachment port

```bash
# no route at all to agent-net:
docker compose exec -T docglean python -c "import urllib.request as u;u.urlopen('http://10.89.0.2:3000/',timeout=4)"      # expect: network unreachable/timeout
# on the shared bridge nothing listens on 3000 (bound to 10.89.0.2 only):
docker compose exec -T docglean python -c "import urllib.request as u;u.urlopen('http://10.89.1.2:3000/mcp',timeout=4)"     # expect: connection refused
# the one thing it can reach, with no ticket => uniform 404:
docker compose exec -T docglean python -c "
import urllib.request as u,urllib.error as e
try: u.urlopen('http://m365:3001/attachment',timeout=4)
except e.HTTPError as x: print(x.code, x.read())"                                                                            # expect: 404 b'Not found'
docker compose exec -T docglean python -c "
import urllib.request as u,urllib.error as e
for p in ('/mcp','/attachment?t=nope&dgk=1&dgx=1&dgs=x'):
  try: u.urlopen('http://m365:3001'+p,timeout=4)
  except e.HTTPError as x: print(p, x.code)"                                                                                    # expect: 404, 404
```

### 7.2 Ticket reuse, spent ticket, expiry

`read-document` hides the URL (it is redacted from any result), so reuse needs a minting tool. Swap in the `negative` profile: same addresses, no proxy, `get-download-url` registered, TTL 20 s.

```bash
docker compose stop m365 && docker compose --profile negative up -d m365-mint
# wait for the log line 'Attachment URLs ... ttl 20s'
mcp '{"jsonrpc":"2.0","id":5,"method":"tools/call","params":{"name":"get-download-url","arguments":{"target":"/me/messages/<MESSAGE_ID>/attachments/<ATTACHMENT_ID>/$value"}}}'
```

Result contains `downloadUrl` like `http://m365:3001/attachment?t=...&dgk=1&dgx=...&dgs=...` (a live credential for 20 s; do not share it). Put it in a shell variable via your terminal only, not chat. `get-download-url` tickets are single-use. Fetch it from the converter side:

```bash
U='<downloadUrl>'
fetch() { docker compose exec -T -e U="$1" docglean python -c "
import os,urllib.request as u,urllib.error as e
try:
  r=u.urlopen(os.environ['U'],timeout=8);print(r.status,r.headers.get('content-type'),len(r.read()))
except e.HTTPError as x: print(x.code,x.read()[:60])"; }
fetch "$U"      # expect: 200 <attachment type> <bytes>
fetch "$U"      # expect: 404 b'Not found'  (spent; identical to unknown/expired)
```

Same through docglean's MCP (bearer from `.env`; run inside the docglean container so `localhost` works):

```bash
# mint a fresh URL first (above), then:
docker compose exec -T -e U="<fresh downloadUrl>" -e T="$(grep ^DOCGLEAN_TOKEN .env | cut -d= -f2)" docglean python -c "
import os,json,urllib.request as u
b=json.dumps({'jsonrpc':'2.0','id':1,'method':'tools/call','params':{'name':'convert_to_markdown','arguments':{'uri':os.environ['U']}}}).encode()
r=u.Request('http://127.0.0.1:8080/mcp',b,{'content-type':'application/json','accept':'application/json, text/event-stream','authorization':'Bearer '+os.environ['T']})
print(u.urlopen(r).read()[:600])"
```

- First run: Markdown. Run it again with the same URL: `fetch_failed` (ticket spent; docglean sees ms-365's 404).
- Expired: mint, wait 25 s, then try both. Direct fetch: 404. docglean: in the keyed mode `blocked_host` (expiry is past its 5 s skew, so the URL is treated as unsigned and never dialled); in the unsigned mode there is no expiry to check, so docglean fetches and gets ms-365's 404, `fetch_failed`.
- Budget of 3 is only visible for `read-document` tickets (`maxRedemptions` 3: docglean may probe then convert). `get-download-url` tickets are 1.
- `get-upload-url` also exists under this profile; ignore it.

Restore: `docker compose --profile negative stop m365-mint && docker compose up -d m365`.

### 7.3 Other cheap checks

- Mode mismatch: set `MS365_MCP_ATTACHMENT_URL_KEY` on ms-365 only (against the unsigned profile above): expect `blocked_host` on every read (not verified), with the operator hint in `message`. Unset it afterwards. (Keyed variant: a different key file on each side gives the same result.)
- `read-document` with `"target":"https://graph.microsoft.com/v1.0/me/messages"`: `invalid_target` — **confirmed**.
- `read-document` with the bare attachment path (no `/$value`): `invalid_target` — **confirmed**.
- Stop docglean (`docker compose stop docglean`), call `read-document`: `proxy_unreachable` after one retry (two warn lines in `docker compose logs m365`). Not exercised in this run.

The network-isolation checks in 7.1 were all run on 2026-10-11 and matched exactly: docglean reaching `10.89.0.2:3000` directly gets "Network is unreachable" (no route to agent-net at all), reaching `10.89.1.2:3000/mcp` on the shared bridge gets "Connection refused" (nothing listens there), and the attachment port gives a uniform 404 both with no ticket at all and with a bogus `t=nope&dgk=1&dgx=1&dgs=x`.

## 8. Teardown

```bash
docker compose --profile negative down -v        # -v deletes the m365-data volume, i.e. the cached login
rm -rf secrets .env
docker rmi m365-mcp:e2e docglean-mcp:e2e
# plain-node path:  docker rm -f docglean-e2e;  node dist/index.js --logout   (from the ms-365 repo)
```

Revoke the app consent in your Microsoft account if this was a throwaway sign-in.

## Appendix A. Keyed variant (signed URLs)

Use it when the docglean build predates the unsigned-ticket change, or you want signatures on top of the address pin. Both sides then read one key from the same file, and the docglean profile must not carry `_UNSIGNED_TICKET_PATH` (a keyless ms-365 would otherwise still be accepted).

```bash
mkdir -p secrets && umask 077
openssl rand -hex 32 > secrets/attachment_url_key
chmod 644 secrets/attachment_url_key     # docglean runs as uid 10001; throwaway test key
docker compose -f docker-compose.yml -f docker-compose.keyed.yml up -d docglean   # Compose 2.24+
```

Use the same `-f` pair for every later `docker compose` command. The override adds the key file and key id to both services, and `MAX_TTL_S` to docglean. A trailing newline in the key file is fine: both sides strip it. Do not let an editor add a BOM. Delete the key at teardown.

### Signing compatibility (checked against code)

| Item              | ms-365 (`src/lib/url-signing.ts`, `attachment-url-config.ts`)                                                     | docglean (`signing.py`)                      | Match                                                                              |
| ----------------- | ----------------------------------------------------------------------------------------------------------------- | -------------------------------------------- | ---------------------------------------------------------------------------------- |
| Canonical string  | `v1\nscheme\nhost\nport\npath\nquery\nexpiry`, port always explicit, query minus `dgk/dgx/dgs`, sorted, re-quoted | same                                         | yes (documented as a port; ms-365 has Python-parity vectors)                       |
| HMAC              | SHA-256, key = UTF-8 bytes, base64url, no padding                                                                 | same                                         | yes                                                                                |
| Key id            | `MS365_MCP_ATTACHMENT_URL_KEY_ID`, default `1`                                                                    | `..._KEY_ID`, default `1`, compared to `dgk` | yes; keep both `1` or change both                                                  |
| TTL               | `MS365_MCP_ATTACHMENT_URL_TTL_S`, default 120, hard max 300                                                       | `..._MAX_TTL_S`, default 300, plus 5 s skew  | yes; ms-365 can never exceed docglean's default. If you raise one, raise the other |
| Key file          | stripped, rejects control chars                                                                                   | stripped                                     | yes                                                                                |
| Host in signature | from `MS365_MCP_ATTACHMENT_URL_BASE` (`m365`), IPv6 literal refused                                               | from `_ORIGIN`                               | yes; both must say `m365:3001`                                                     |

No blocking inconsistency found. Two operator-side facts that do bite: the docglean `_ORIGIN` must equal the URL base exactly (scheme, host, port), and `_ALLOW_CIDRS` must contain the resolved address of that host (here `10.89.1.2/32`).
