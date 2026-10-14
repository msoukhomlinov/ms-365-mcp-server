# Teams bot service: execution log

Ledger from the subagent-driven execution of `2026-10-14-teams-bot-service.md` (Tasks 1-11 plus a final fix wave). The service lives in the local repo `/Users/maxs/gitrepos/teams-bot-mcp` (25 commits on `main`, 158 tests, not pushed anywhere). Task 12 (reconcile with the Phase 0 tenant probe) is not done: it needs the user to run the probe. Lines marked `minor (deferred)` and the open items at the end are the work list for Task 12 and later.

```text
# SDD ledger — plan: /Users/maxs/gitrepos/ms-365-mcp-server/docs/superpowers/plans/2026-10-14-teams-bot-service.md
Spec read: docs/superpowers/specs/2026-10-14-teams-bot-service-design.md (reachable, in the fork checkout).
Repo: /Users/maxs/gitrepos/teams-bot-mcp (local only, no GitHub repo, no push). Initial empty commit d37e274.

## Pre-flight scan
| Pair / task | Produces vs consumes | Finding |
|---|---|---|
| T1 -> T2 | helpers.testConfig returns Config (T2 type); T1 does not run typecheck | consistent; T1 runs only vitest on errors.test.ts |
| T2 -> T3 | Config/Team/Channel | consistent |
| T2 -> T9, T10, helpers | Config fields tenantId, appId, secret, token, serviceUrl, teams | consistent |
| T3 -> T9 | resolveTeam, resolveChannel, requireWrite, assertNumericId | consistent |
| T4 -> T5, T8, T10 | TokenProvider.get(scope); helpers.fakeTokens has the same shape | consistent |
| T5 -> T9 | GraphClient 4 methods | consistent |
| T6 -> T9 | toMessage(raw, appId) | consistent |
| T7 -> T8, T9 | Entity, convertMentions | consistent |
| T8 -> T9 | ConnectorClient post/reply/edit/remove | consistent |
| T9 -> T10 | Tool, buildTools, AuditLine | consistent |
| T11 self | config.example.json read by test relative to cwd | consistent |
| T12 | gated on the user's probe results | not executable by a subagent; stop before it |
Self-consistency per task: T1-T11 tests agree with the code they specify (checked by the plan self-review); defects found below.

## Preflight rulings
Ruling: eslint recommended flags unused args, and helpers.fakeTokens has a sole unused arg _scope — Task 1 adds rule '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_' }] to eslint.config.js — cost if wrong: none
Ruling: plan lacks ignore files, and format:check / lint would scan .superpowers workspace files and package-lock.json — Task 1 adds .prettierignore (.superpowers, dist, node_modules, package-lock.json) and adds '.superpowers' to the eslint ignores — cost if wrong: none
Ruling: plan code is not guaranteed prettier-clean at printWidth 100 — every implementer runs 'npx prettier --write' on the files it touched before committing — cost if wrong: a formatting-only follow-up commit
Ruling: model tiers — implementers transcribe complete code: haiku for T1,T2,T3,T6,T7; sonnet for T4,T5,T8,T9,T10,T11 (security-critical or integration); reviewers sonnet, opus for T5,T8,T9,T10; final review fable — cost if wrong: extra fix rounds
Ruling: Task 12 is gated on the user running the Phase 0 tenant probe; execution stops after Task 11 and the final review — cost if wrong: none
Task 1: minor (deferred): npm warns esbuild postinstall not covered by allowScripts (check at build in Task 10); npm audit reports 1 low-severity vulnerability (investigate at final review); typecheck deferred to Task 2 as expected
Task 1: minor (deferred): eslint.config.js uses tseslint.config(), deprecated in newer typescript-eslint 8.x (cosmetic)
Task 1: complete (commits d37e274..052ebf2, review clean)
Task 2: first reviewer dispatch returned no verdict (no file tools); re-dispatched, not counted as a review
Task 2: minor (deferred): serviceUrl path is unconstrained (e.g. /teams/../x normalises to /x/, //evil.com/ stays on host smba.trafficmanager.net) — host is always pinned so the token cannot go elsewhere; pin pathname to /teams/ at final review if wanted
Task 2: review found 1 Important (serviceUrl accepts query/hash, base URL breaks at send time) — entering fix round 1/5 (resumed implementer)
Task 2: fix round 1/5 (1 addressed, 0 open; commits 2019ad3..56a4be3)
Task 2: complete (commits 052ebf2..56a4be3, review clean after 1 fix round)
Task 3: minor (deferred): assertNumericId tests do not pin "1\n", unicode digits or array inputs (verified rejecting today by probe)
Task 3: minor (deferred): assertNumericId error text says "message id" even for thread (cosmetic)
Task 3: complete (commits 56a4be3..5ab8d91, review clean)
Task 4: Ruling: plan test pattern `.catch((e) => e as ToolError)` fails typecheck (TS2339, type is string | ToolError); implementer changed it to `(await x.catch((e: unknown) => e)) as ToolError`, assertions unchanged — the same pattern exists in the plan tests for Tasks 5 and later, so those implementers are told to use the cast form — cost if wrong: none, test typing only
Task 4: minor (deferred): token.ts echoes data.error uncapped (optional allowlist /^[a-z_]{1,64}$/); expires_in 0 or Infinity handling (use Number.isFinite && > 0); a 2xx without access_token or 5xx is worded as "client secret may be wrong"; catch discards cause with no sanitized log; no tests for timeout or expiry edges
Task 4: complete (commits 5ab8d91..981d9c4, review clean)
Task 5: minor (deferred): validateCursor runs its raw-string regex before the URL parser strips tab/newline, and misses %2e%2e; accepted variants normalise to the expected path or a nonexistent one (no other resource reached). Fix: reject [\x00-\x20\x7f] in the cursor and check u.pathname; return u.href so the fetched URL is the validated one
Task 5: minor (deferred): the startsWith(GRAPH + '/') guard in get() is weaker than its error suggests (https://graph.microsoft.com/v1.0/%2e%2e/beta/teams passes it); no current caller can produce it. Fix: check new URL(url).origin and pathname prefix, fetch u.href
Task 5: minor (deferred): cursors with uppercase scheme/host or newline pass validateCursor but fail startsWith, so they come back as 'internal' not 'validation' (classification only)
Task 5: minor (deferred): cursor query string is unvalidated: ?$expand=replies&$top=999 accepted on a valid messages URL; stays inside the allowlisted team/channel and Graph enforces RSC — a quirk, not a boundary crossing
Task 5: minor (deferred): add redirect: 'error' to Graph fetch (undici already strips Authorization on cross-origin redirects, probed)
Task 5: minor (deferred): a 2xx non-JSON body makes res.json() throw a SyntaxError containing a body fragment; wrap into ToolError('unavailable') (tool layer maps non-ToolError to 'internal' so it does not reach the agent today); error.code copied uncapped
Task 5: minor (deferred): listMembers stops silently after 5 pages (caller cannot tell); throw or return a truncated flag
Task 5: minor (deferred): retried response bodies are not cancelled; a non-ToolError token failure is retried and reported as 'could not be reached'
Task 5: minor (deferred): Graph client relies on Task 9's ^\d{1,20}$ check for thread ids (reviewer probe: a replies cursor for hostile thread '1/../../../../users' passes validateCursor); consider validating ids inside the client too
Task 5: complete (commits 981d9c4..7c7f44f, review clean)
Task 6: review found 2 Important — (1) isOwn strict === with a config appId that may be uppercase, fixed in round 1; (2) html-to-text defaults pass hidden-style text, img alt text and URLs
Task 6: Ruling: html-to-text default rendering stays as written (wordwrap false) — there is no built-in display:none filter, the spec already treats message text as untrusted data and labels it so in the tool descriptions, and keeping img src and link hrefs matches the spec's "hosted-content references are kept as Graph paths" — cost if wrong: hidden-style injection text reaches the agent, labeled untrusted; revisit at final review if wanted
Task 6: minor (deferred): body.contentType compare is case-sensitive ('HTML' returns raw markup); createdDateTime unguarded; from.device maps to kind system undocumented; tests do not pin link/img rendering or malformed shapes
Task 6: entering fix round 1/5 (resumed implementer) for finding (1)
Task 6: fix round 1/5 (1 addressed, 0 open; commits 28a67b0..b98214b)
Task 6: complete (commits 7c7f44f..b98214b, review clean after 1 fix round)
Task 7: minor (deferred, include in the final fix wave): ATTEMPT regex /<at\b[^>]*\buser-id\b/i is quadratic on many '<at ' with no '>' (measured 2.3 s at 100k chars; ~1.5 s at the 80,000-char text cap, event-loop stall from a bearer-authenticated caller) — bound it with [^>]{0,500}
Task 7: minor (deferred): empty or whitespace-only display name (<at user-id="GUID"></at>) is accepted, producing <at></at> and mentioned.name "" — reject empty-after-trim names
Task 7: minor (deferred): hand-built <at>Jane</at> / <at id="0">Jane</at> without user-id pass through with no entity (no notification) — say in the tool description that mentions must use the shorthand; text outside tags intentionally not escaped; tests do not pin newline/tab, self-closing, attribute order, $& name or empty name
Task 7: open question for the probe: whether Teams matches entity.text (XML-escaped label) against the escaped body for names containing & " > ' — covered by the Phase 0 mention probe (Task 12)
Task 7: complete (commits b98214b..1805731, review clean)
Task 8: Ruling: plan test "never sends the token anywhere but the configured service URL" used json({ id: "1" }) as the fake post response, but the plan's post throws unknown when activityId is missing; implementer changed only the fake to json({ id: "1", activityId: "1" }), assertion and code unchanged — cost if wrong: none, test fixture only
Task 8: minor (deferred, include in the final fix wave): fetch follows redirects, so a 307 on a write re-POSTs the message body to another host (Authorization is stripped by undici, probed) — add redirect: 'error' so a redirect becomes unknown
Task 8: minor (deferred, include in the final fix wave): reply() returns messageId '' when the response has no id or fails to parse — copy post's guard (throw unknown)
Task 8: minor (deferred): 424 (and arguably 408/409/412) reported as a definite 'validation' rejection; the activity body shape is written inline in 3 places and the post envelope inline once — an activity(text, entities) helper would make the probe-driven change one edit; client has no own ^\d{1,20}$ id check (encodeURIComponent leaves '..' as is); shared UNKNOWN singleton error instance (stack points at module load); 5xx bodies not cancelled
Task 8: minor (deferred): tests are weak in places — token-only-to-serviceUrl test can only pass by construction and does not check the auth header per call or redirects; the "401 and 403" test covers only 403; no timeout, bad-200-body, token-error pass-through, sleep-count or PUT-body assertions
Task 8: complete (commits 1805731..02f4a91, review clean)
Task 9: first implementer dispatch BLOCKED (no file/shell tools loaded, nothing written, no commits); re-dispatched with a tool-search hint, same model, same BASE 02f4a91
Task 9: minor (deferred, include in the final fix wave): the 80,000-char text limit is checked before convertMentions, and XML-escaping can grow the text ~6x (probe: 79,000 apostrophes in a mention name reached connector.post as 474,009 chars; the Connector's 413 maps to validation) — re-check conv.text.length after conversion
Task 9: minor (deferred, include in the final fix wave): audit(... 'ok') and the return sit inside the write's try block, so an audit sink that throws after a successful write is reported to the agent as 'internal' and audited twice, inviting a duplicate retry — move the ok-audit/return out of the try or wrap audit in its own try
Task 9: minor (deferred): wrong-type arguments (e.g. cursor as an array) are classed 'internal' not 'validation' (production zod rejects them earlier); list-members ignores a non-string name
Task 9: minor (deferred): tests weaker than their names — "rejects thread before any request" checks only getMessage/reply (not listReplies/edit/remove), the read-only test does not assert reply/edit were uncalled, no test for non-string/invalid cursors, audit content on unknown-alias error, or text length after conversion
Task 9: complete (commits 02f4a91..a5a3ff1, review clean)
Task 10: Ruling: plan http test passed a headers array literal that fails typecheck (TS2345); implementer hoisted it into a typed const, assertions unchanged — cost if wrong: none
Task 10: npm audit: 1 low severity (esbuild dev server, GHSA-g7r4-m6w7-qqqr, Windows file read, dev dependency only) — nothing fixed; esbuild postinstall warning did not reproduce, build works
Task 10: minor (deferred, include in the final fix wave): no minimum bearer-token length (a 1-character token file is accepted) — require >= 32 chars in config; TEAMS_BOT_PORT validation ("" binds a random port, "abc" throws ERR_SOCKET_BAD_PORT) — validate an integer 1-65535; `void transport.close(); void mcp.close();` have no .catch (Node 22 unhandled rejection would kill the process); no end-to-end test of tools/list and tools/call over HTTP (reviewer's probe passed: /private/tmp/claude-501/-Users-maxs-gitrepos-ms-365-mcp-server/c7515d86-1988-433e-b2cc-ba335ffcf0ba/scratchpad/probe/probe.ts) — add one
Task 10: minor (deferred): /readyz stampede (20 concurrent calls = 20 ready() calls; token provider single-flight bounds the real work) and a failure cached 30 s; no SIGTERM handler (matters if PID 1 in a container — Task 11 image uses node directly as CMD, consider an init/handler); consider requestTimeout ~30 s if bound off loopback; Origin match is exact so a trailing-slash entry never matches (document)
Task 10: complete (commits a5a3ff1..4e2585b, review clean)
Task 11: Docker daemon was not running on this machine, so the image was NOT built and the docker run exit-1 check was NOT done — unverified, carry into the final message; /healthz returns body "ok" (design says 200 and nothing else; harmless wording gap); SECURITY.md names no private reporting channel (open item with the repo name)
Task 11: harness security notice flagged secret exposure via docker-compose.example.yml; verified by reading: compose uses ./secrets/teams_bot_secret, ./secrets/teams_bot_token and ./config.json, but .gitignore (node_modules, dist, .env, *.log) and .dockerignore (node_modules, dist, .git, test) exclude neither secrets/ nor config.json, so a user following the example can commit credentials and the build stage COPY . . pulls them into the build context (final stage copies only dist/) — real finding, enters Task 11 fix round: add secrets/ and config.json to both ignore files and a README note
Task 11: review approved the docs (all README/SECURITY claims match src and spec; honest about the unrun probe) with Minor findings only; I re-graded the missing secrets/ and config.json ignore entries as Important by effect (credential commit risk) — entering fix round 1/5 (resumed implementer)
Task 11: minor (deferred): README /readyz line says it "fetches both tokens (cached 30 seconds)" — it is the readiness result that is cached; SECURITY.md names no private reporting channel (open item with the repo name)
Task 11: fix round 1/5 (1 addressed, 0 open; commits 5e40adc..7bbbb54)
Task 11: complete (commits 4e2585b..7bbbb54, review clean after 1 fix round)
Task 12: not executable by a subagent (needs the user to run the Phase 0 tenant probe and supply results and redacted fixtures); execution stops after the final review
Final: whole-branch review (fable): no Critical; Review Focus 1-5 verified on the real server (46-check e2e driver at /private/tmp/claude-501/-Users-maxs-gitrepos-ms-365-mcp-server/c7515d86-1988-433e-b2cc-ba335ffcf0ba/scratchpad/e2e/driver.mjs). Fix list of 14 items written to final-fix-list.md; one fix dispatch.
Final: Ruling: do NOT pin serviceUrl pathname to /teams/ (the Task 2 deferred idea) — it would break Task 12's regional fallback (/amer/, /emea/, /apac/ share the host); the host pin is the control — cost if wrong: a path-prefix wobble inside smba.trafficmanager.net only
Final: Ruling: bearer token minimum length 32 characters is added (config validation) — a placeholder such as "changeme" in tokenFile is otherwise accepted on a listener that can be exposed — cost if wrong: operators with shorter tokens must regenerate them
Final: open (cannot be fixed by the wave): spec item "one opt-in live test" (needs the probe team), the Teams manifest file, saving a Graph read of a bot-authored message as a fixture, root-vs-reply conversationId flag — all fold into Task 12 after the user runs the probe; Docker image build unverified unless the daemon is running (fix-list item 14)
Final: minor (deferred, recommended to leave by the final reviewer): schema-level rejections not audited; worst-case read-thread latency ~130 s; duplicated Retry-After arithmetic; token 5xx wording; no GHCR workflow until a remote exists; listMembers 5-page cap; Task 1-8 items listed earlier not in the fix wave
Final: fix wave landed as 9 commits 7bbbb54..fa5cd1b (155 tests green); Docker item 14 skipped (daemon down); scoped re-review (opus) dispatched
Final: harness security notice flagged a parser-differential bypass in src/mentions.ts; verified by running the real regexes: after fix-wave item 5 bounded ATTEMPT to /<at\b[^>]{0,500}\buser-id\b/i, a malformed tag with user-id more than 500 chars after <at (extra attribute `<at data-x="<600 chars>" user-id="GUID">Jane</at>`, or single quotes after 600 spaces) matches neither TAG nor ATTEMPT and is posted as literal text — a regression introduced by the fix list's own item 5 and a breach of the spec's "a literal tag is never sent"
Final: Ruling: one narrow correction round is dispatched even though the final-review rule says no second fix wave, because the defect was introduced by the wave itself and breaks a spec property; the correct fix is a LINEAR scan (find each '<at' with /<at\b/gi, take the text to the next '>' or end of string, skip occurrences already inside a checked region, test /\buser-id\b/i on that span) rather than another bound — cost if wrong: one extra commit and one extra re-check
Final: fixed bounded-regex bypass (spec "literal tag never sent") — 4 new tests RED→GREEN in test/mentions.test.ts (long-gap valid tag converts; invalid GUID, single-quote/extra-attribute, newline-gap variants rejected; 200,000 x '<at ' under 100 ms), commit 29bb1fe, suite 158/158; scoped re-check (sonnet) ADDRESSED, worst adversarial input up to ~500k chars took 10 ms
Final: minor (deferred): a '>' inside an attribute value before user-id (`<at title="a>b" user-id="guid">`) is not detected and passes through as a literal tag — pre-existing edge case, not a regression; entities[].text labels are not counted in the post-conversion length check
Final: all fix-wave items 1-13 addressed and re-reviewed; item 14 (Docker build, run without config exits 1, non-root uid) NOT verified: daemon down
```
