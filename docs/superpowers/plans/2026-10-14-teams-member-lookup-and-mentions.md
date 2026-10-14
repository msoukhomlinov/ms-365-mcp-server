# Teams member lookup and mention shorthand Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an optional `name` filter to the two Teams member-list tools (PR A) and an inline `<at user-id="...">` mention shorthand to the seven Teams message tools (PR B), each as an independent upstream PR.

**Architecture:** Both changes live in `executeGraphTool` in `src/graph-tools.ts`, next to the existing request/response rewriters (`normalizeSearchQueryParam`, the `body` repair block, the `$select` projection). Each adds one small exported pure helper (unit-testable), one wiring site, a schema/param description, and `llmTip` updates in `src/endpoints.json`.

**Tech Stack:** TypeScript, zod, vitest (`npm test`), prettier, eslint, tsup.

**Spec:** `docs/superpowers/specs/2026-10-14-teams-member-lookup-and-mentions-design.md` (in the main checkout on `integration/hermes-family`; it is not part of either PR).

## Global Constraints

- Each PR branches from `upstream/main`, ships on its own, and does not depend on the other.
- No existing tool name, required parameter or correct request body changes behavior.
- Follow upstream patterns: `endpoints.json` entries and `llmTip`, request rewriting in `executeGraphTool` in `src/graph-tools.ts`, tests in `src/__tests__/graph-tools.test.ts`.
- Tests are written first and must be seen failing.
- The spec file stays on the fork integration branch and is not part of either PR.
- No AI attribution anywhere: commits, PR titles/bodies, comments, code comments. No `Co-Authored-By` trailers, no "Generated with" lines. If you dispatch a subagent, include this rule in its brief.
- Do not push or open a PR without asking the user first.

## Review Focus

1. `name` that is empty or whitespace-only: return the full list, not an empty one (Task 2 test).
2. Members with `null` or missing `displayName`/`email`: no crash, and no match (Task 2 test).
3. Error or non-JSON response with `name` set: passed through untouched (Task 2 test).
4. Raw string body, or `contentType` missing or `text`, on a message tool: untouched, no throw (Task 4 tests).
5. Same user mentioned twice, and shorthand mixed with a hand-built `<at id="0">`: ids unique, every tag has a matching `mentions` entry (Task 3 tests).

---

## Setup (once per PR)

Use a worktree so the dirty fork checkout is untouched and the branch has no fork-only files.

```bash
cd /Users/maxs/gitrepos/ms-365-mcp-server
git fetch upstream
git worktree add ../ms-365-mcp-server-pr-a -b feat/teams-member-name-filter upstream/main
git worktree add ../ms-365-mcp-server-pr-b -b feat/teams-mention-shorthand upstream/main
(cd ../ms-365-mcp-server-pr-a && npm ci)
(cd ../ms-365-mcp-server-pr-b && npm ci)
```

Line numbers below are approximate. Locate each site by its anchor text.

---

# Part A: `name` filter on member lists (worktree `pr-a`)

### Task 1: `filterMembersByName` helper

**Files:**

- Modify: `src/graph-tools.ts` (add near `normalizeSearchQueryParam`)
- Test: `src/__tests__/graph-tools.test.ts`

**Interfaces:**

- Produces: `export function filterMembersByName(parsed: unknown, name: string): unknown` and `const MEMBER_NAME_FILTER_TOOLS: Set<string>`, `const MEMBER_NAME_PARAM_DESCRIPTION: string`.

- [ ] **Step 1: Write the failing tests**

Append a new describe at the end of the top-level `describe` in `src/__tests__/graph-tools.test.ts`:

```ts
describe('filterMembersByName', () => {
  const body = {
    value: [
      { id: '1', displayName: 'Jane Doe', email: 'jane@contoso.com' },
      { id: '2', displayName: 'John Smith', email: 'jsmith@contoso.com' },
      { id: '3', displayName: null, email: null },
    ],
    '@odata.count': 3,
  };
  const ids = (r: any) => r.value.map((m: any) => m.id);

  it('matches displayName case-insensitively', async () => {
    const { filterMembersByName } = await loadModule();
    expect(ids(filterMembersByName(body, 'JANE'))).toEqual(['1']);
  });

  it('matches email', async () => {
    const { filterMembersByName } = await loadModule();
    expect(ids(filterMembersByName(body, 'jsmith@'))).toEqual(['2']);
  });

  it('returns an empty value array when nothing matches', async () => {
    const { filterMembersByName } = await loadModule();
    expect(ids(filterMembersByName(body, 'zzz'))).toEqual([]);
  });

  it('returns everything for an empty or whitespace-only name', async () => {
    const { filterMembersByName } = await loadModule();
    expect(ids(filterMembersByName(body, '   '))).toEqual(['1', '2', '3']);
  });

  it('does not match or throw on null fields', async () => {
    const { filterMembersByName } = await loadModule();
    expect(ids(filterMembersByName(body, 'null'))).toEqual([]);
  });

  it('keeps other top-level fields', async () => {
    const { filterMembersByName } = await loadModule();
    expect((filterMembersByName(body, 'jane') as any)['@odata.count']).toBe(3);
  });

  it('returns non-collection input unchanged', async () => {
    const { filterMembersByName } = await loadModule();
    expect(filterMembersByName('plain', 'x')).toBe('plain');
    expect(filterMembersByName({ id: '1' }, 'x')).toEqual({ id: '1' });
  });
});
```

- [ ] **Step 2: Run to verify failure**

Run: `npx vitest run src/__tests__/graph-tools.test.ts -t filterMembersByName`
Expected: FAIL with `filterMembersByName is not a function`.

- [ ] **Step 3: Implement**

Add to `src/graph-tools.ts` just above the `normalizeSearchQueryParam` doc comment (reuse the file's existing `isPlainObject`):

```ts
const MEMBER_NAME_FILTER_TOOLS = new Set(['list-team-members', 'list-chat-members']);

const MEMBER_NAME_PARAM_DESCRIPTION =
  "Case-insensitive text to match against each member's displayName and email. Applied after $select projection, so only fields you can see are matched. Without fetchAllPages only the fetched page is filtered, and @odata.count still reports the unfiltered total.";

/**
 * Keeps only the members whose displayName or email contains `name`. Runs on the final
 * response body, after $select projection, so a field the caller cannot see is never
 * matched on. Blank `name` and non-collection bodies pass through unchanged.
 */
export function filterMembersByName(parsed: unknown, name: string): unknown {
  const needle = name.trim().toLowerCase();
  if (needle === '' || !isPlainObject(parsed) || !Array.isArray(parsed.value)) return parsed;
  const matches = (member: unknown): boolean =>
    isPlainObject(member) &&
    ['displayName', 'email'].some((field) => {
      const value = member[field];
      return typeof value === 'string' && value.toLowerCase().includes(needle);
    });
  return { ...parsed, value: parsed.value.filter(matches) };
}
```

- [ ] **Step 4: Run to verify pass**

Run: `npx vitest run src/__tests__/graph-tools.test.ts -t filterMembersByName`
Expected: PASS (7 tests).

- [ ] **Step 5: Commit**

```bash
git add src/graph-tools.ts src/__tests__/graph-tools.test.ts
git commit -m "feat(tools): add member name filter helper"
```

### Task 2: Wire `name` into the member-list tools

**Files:**

- Modify: `src/graph-tools.ts` (four sites, below)
- Modify: `src/endpoints.json` (`list-team-members`, `list-chat-members` tips)
- Test: `src/__tests__/graph-tools.test.ts`

**Interfaces:**

- Consumes: `MEMBER_NAME_FILTER_TOOLS`, `MEMBER_NAME_PARAM_DESCRIPTION`, `filterMembersByName` from Task 1.

- [ ] **Step 1: Write the failing tests**

Append inside the top-level `describe`:

```ts
describe('member name filter (tool wiring)', () => {
  const membersBody = {
    value: [
      { id: '1', displayName: 'Jane Doe', email: 'jane@contoso.com' },
      { id: '2', displayName: 'John Smith', email: 'jsmith@contoso.com' },
      { id: '3', displayName: null, email: null },
    ],
    '@odata.count': 3,
  };
  const page = (b: unknown) => ({ content: [{ type: 'text', text: JSON.stringify(b) }] });
  const ids = (result: any) => JSON.parse(result.content[0].text).value.map((m: any) => m.id);

  async function run(
    args: Record<string, unknown>,
    responses: any[] = [page(membersBody)],
    alias = 'list-team-members',
    path = '/teams/{team-id}/members'
  ) {
    mockEndpoints.push(makeEndpoint({ alias, path }));
    mockEndpointsJson = [makeConfig({ toolName: alias, pathPattern: path })];
    const graphClient = createMockGraphClient(responses);
    const server = createMockServer();
    const { registerGraphTools } = await loadModule();
    registerGraphTools(server as any, graphClient as any);
    const result = await server.tools.get(alias)!.handler({ 'team-id': 't1', ...args });
    return { result, graphClient };
  }

  it('filters list-team-members by name', async () => {
    const { result } = await run({ name: 'jane' });
    expect(ids(result)).toEqual(['1']);
  });

  it('filters list-chat-members by email', async () => {
    const { result } = await run(
      { 'chat-id': 'c1', name: 'jsmith' },
      [page(membersBody)],
      'list-chat-members',
      '/chats/{chat-id}/members'
    );
    expect(ids(result)).toEqual(['2']);
  });

  it('does not send name to Graph', async () => {
    const { graphClient } = await run({ name: 'jane' });
    expect(graphClient.graphRequest.mock.calls[0][0]).not.toContain('name');
  });

  it('returns the full list for a whitespace-only name', async () => {
    const { result } = await run({ name: '  ' });
    expect(ids(result)).toEqual(['1', '2', '3']);
  });

  it('leaves an error response untouched', async () => {
    const err = { content: [{ type: 'text', text: '{"error":"nope"}' }], isError: true };
    const { result } = await run({ name: 'jane' }, [err]);
    expect(result.content[0].text).toBe('{"error":"nope"}');
  });

  it('keeps @odata.count as the unfiltered total', async () => {
    const { result } = await run({ name: 'jane' });
    expect(JSON.parse(result.content[0].text)['@odata.count']).toBe(3);
  });

  it('filters after fetchAllPages merges the pages', async () => {
    const { result } = await run({ name: 'ann', fetchAllPages: true }, [
      page({
        value: [{ id: '1', displayName: 'Ann', email: 'a@x.com' }],
        '@odata.nextLink': 'https://graph.microsoft.com/v1.0/teams/t1/members?$skiptoken=x',
      }),
      page({
        value: [
          { id: '2', displayName: 'Bob', email: 'b@x.com' },
          { id: '3', displayName: 'Annette', email: 'n@x.com' },
        ],
      }),
    ]);
    expect(ids(result)).toEqual(['1', '3']);
  });

  it('matches only fields that survive $select projection', async () => {
    const { result } = await run({ select: 'id,displayName', name: 'contoso' });
    expect(ids(result)).toEqual([]);
  });

  it('still matches a projected field', async () => {
    const { result } = await run({ select: 'id,displayName', name: 'jane' });
    expect(JSON.parse(result.content[0].text).value).toEqual([
      { id: '1', displayName: 'Jane Doe' },
    ]);
  });
});
```

- [ ] **Step 2: Run to verify failure**

Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "member name filter"`
Expected: FAIL (the filter is not wired; `name` is dropped, all members return).

- [ ] **Step 3: Implement the wiring**

In `src/graph-tools.ts`, four edits:

1. In the parameter loop, immediately before the `// Skip control parameters` block:

```ts
if (paramName === 'name' && MEMBER_NAME_FILTER_TOOLS.has(tool.alias)) continue;
```

2. Before the loop (next to `const strayBodyFields`):

```ts
const memberNameFilter =
  MEMBER_NAME_FILTER_TOOLS.has(tool.alias) && typeof params.name === 'string'
    ? params.name
    : undefined;
```

3. Where `forceJsonOutput` is decided, replace the existing condition:

```ts
const willFilterMembers =
  memberNameFilter !== undefined &&
  memberNameFilter.trim() !== '' &&
  params.excludeResponse !== true;
if (mergePages || willProject || willProjectBatch || willFilterMembers) {
  options.forceJsonOutput = true;
}
```

4. After the `willProjectBatch` block and before `if (response?.content?.[0]?.text) {` that logs `Response size`:

```ts
if (willFilterMembers && !response?.isError && response?.content?.[0]?.text) {
  try {
    const filtered = filterMembersByName(
      JSON.parse(response.content[0].text),
      memberNameFilter as string
    );
    response.content[0].text = graphClient.serialize(filtered);
    if (isPlainObject(filtered) && Array.isArray(filtered.value) && response._meta) {
      response._meta = { ...response._meta, result_count: filtered.value.length };
    }
  } catch {
    // Body was not JSON after all; nothing to filter.
  }
}
```

5. In the `paramSchema` construction, directly after the `isFetchAllPagesApplicable` block:

```ts
if (MEMBER_NAME_FILTER_TOOLS.has(tool.alias)) {
  paramSchema['name'] = z.string().describe(MEMBER_NAME_PARAM_DESCRIPTION).optional();
}
```

In `src/endpoints.json`:

- `list-team-members` (no tip today): add `"llmTip": "Lists members of a team. Each member has displayName, email, roles and userId. Pass name to keep only members whose displayName or email contains that text (case-insensitive); use the returned userId in an @mention."`
- `list-chat-members`: append to the existing tip: ` Pass name to keep only members whose displayName or email contains that text (case-insensitive); use the returned userId in an @mention.`

- [ ] **Step 4: Run to verify pass, then the whole suite**

Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "member name filter"` then `npm test`
Expected: PASS, no regressions.

- [ ] **Step 5: Verify and commit**

Run: `npx prettier --write src/graph-tools.ts src/endpoints.json src/__tests__/graph-tools.test.ts && npm run verify`
Expected: all green. If `npm run generate` changes files under `src/generated`, include them.

```bash
git add -A
git commit -m "feat(tools): add name filter to list-team-members and list-chat-members"
```

- [ ] **Step 6: STOP and ask the user** before `git push` / opening the PR. PR text: what, why, tests. No AI attribution.

---

# Part B: mention shorthand (worktree `pr-b`)

### Task 3: `rewriteMentionShorthand` helper

**Files:**

- Modify: `src/graph-tools.ts` (add near `bodyForRecipientWalk`)
- Test: `src/__tests__/graph-tools.test.ts`

**Interfaces:**

- Produces: `export function rewriteMentionShorthand(body: unknown): { body: unknown; error?: string }` and `const MENTION_SHORTHAND_TOOLS: Set<string>` (the seven aliases).

- [ ] **Step 1: Write the failing tests**

```ts
describe('rewriteMentionShorthand', () => {
  const U1 = '11111111-2222-3333-4444-555555555555';
  const U2 = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
  const html = (content: string, extra: Record<string, unknown> = {}) => ({
    body: { contentType: 'html', content },
    ...extra,
  });
  const mention = (id: number, userId: string, name: string) => ({
    id,
    mentionText: name,
    mentioned: { user: { id: userId, displayName: name, userIdentityType: 'aadUser' } },
  });

  it('rewrites one tag and builds the mentions array', async () => {
    const { rewriteMentionShorthand } = await loadModule();
    expect(rewriteMentionShorthand(html(`Hi <at user-id="${U1}">Jane Doe</at>!`)).body).toEqual(
      html('Hi <at id="0">Jane Doe</at>!', { mentions: [mention(0, U1, 'Jane Doe')] })
    );
  });

  it('numbers several tags in order, including a repeated user', async () => {
    const { rewriteMentionShorthand } = await loadModule();
    const out = rewriteMentionShorthand(
      html(
        `<at user-id="${U1}">Jane</at> <at user-id="${U2}">Bob</at> <at user-id="${U1}">Jane</at>`
      )
    ).body as any;
    expect(out.body.content).toBe('<at id="0">Jane</at> <at id="1">Bob</at> <at id="2">Jane</at>');
    expect(out.mentions.map((m: any) => m.id)).toEqual([0, 1, 2]);
  });

  it('starts above ids already in the content and mentions', async () => {
    const { rewriteMentionShorthand } = await loadModule();
    const input = html(`<at id="0">Old</at> <at user-id="${U1}">Jane</at>`, {
      mentions: [mention(0, U2, 'Old')],
    });
    const out = rewriteMentionShorthand(input).body as any;
    expect(out.body.content).toBe('<at id="0">Old</at> <at id="1">Jane</at>');
    expect(out.mentions).toEqual([mention(0, U2, 'Old'), mention(1, U1, 'Jane')]);
  });

  it('passes a correct hand-built body through unchanged', async () => {
    const { rewriteMentionShorthand } = await loadModule();
    const input = html('<at id="0">Jane</at>', { mentions: [mention(0, U1, 'Jane')] });
    expect(rewriteMentionShorthand(input).body).toBe(input);
  });

  it('returns an error for a user-id that is not a GUID', async () => {
    const { rewriteMentionShorthand } = await loadModule();
    const input = html('<at user-id="jane@contoso.com">Jane</at>');
    const out = rewriteMentionShorthand(input);
    expect(out.error).toContain('jane@contoso.com');
    expect(out.body).toBe(input);
  });

  it('leaves non-html content alone', async () => {
    const { rewriteMentionShorthand } = await loadModule();
    const input = { body: { contentType: 'text', content: `<at user-id="${U1}">Jane</at>` } };
    expect(rewriteMentionShorthand(input).body).toBe(input);
  });

  it('leaves a body with no contentType alone', async () => {
    const { rewriteMentionShorthand } = await loadModule();
    const input = { body: { content: `<at user-id="${U1}">Jane</at>` } };
    expect(rewriteMentionShorthand(input).body).toBe(input);
  });

  it('leaves a raw string body alone', async () => {
    const { rewriteMentionShorthand } = await loadModule();
    expect(rewriteMentionShorthand('{"body":{}}').body).toBe('{"body":{}}');
    expect(rewriteMentionShorthand(null).body).toBeNull();
  });

  it('keeps the display name verbatim', async () => {
    const { rewriteMentionShorthand } = await loadModule();
    const out = rewriteMentionShorthand(html(`<at user-id="${U1}">Smith &amp; Co</at>`))
      .body as any;
    expect(out.body.content).toBe('<at id="0">Smith &amp; Co</at>');
    expect(out.mentions[0].mentionText).toBe('Smith &amp; Co');
  });
});
```

- [ ] **Step 2: Run to verify failure**

Run: `npx vitest run src/__tests__/graph-tools.test.ts -t rewriteMentionShorthand`
Expected: FAIL with `rewriteMentionShorthand is not a function`.

- [ ] **Step 3: Implement**

Add to `src/graph-tools.ts` above `bodyForRecipientWalk`:

```ts
const MENTION_SHORTHAND_TOOLS = new Set([
  'send-channel-message',
  'reply-to-channel-message',
  'update-channel-message',
  'update-channel-message-reply',
  'send-chat-message',
  'reply-to-chat-message',
  'update-chat-message',
]);

const AAD_GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const SHORTHAND_AT = /<at\s+user-id="([^"]*)">([^<]*)<\/at>/gi;

/**
 * Teams mentions need numbered <at id="N"> tags plus a matching `mentions` array, and a
 * mismatch is a 400. Models write <at user-id="<aad-guid>">Name</at> instead; this
 * numbers each tag above any id already present and builds the mentions entries.
 * Bodies without the shorthand, non-html content and non-object bodies pass through.
 */
export function rewriteMentionShorthand(body: unknown): { body: unknown; error?: string } {
  if (!isPlainObject(body) || !isPlainObject(body.body)) return { body };
  const itemBody = body.body;
  const content = itemBody.content;
  if (typeof content !== 'string' || String(itemBody.contentType).toLowerCase() !== 'html') {
    return { body };
  }

  const existing = Array.isArray(body.mentions) ? body.mentions : [];
  let nextId = 0;
  for (const m of existing) {
    if (isPlainObject(m) && typeof m.id === 'number') nextId = Math.max(nextId, m.id + 1);
  }
  for (const m of content.matchAll(/<at\s+id="(\d+)"/gi)) {
    nextId = Math.max(nextId, Number(m[1]) + 1);
  }

  const added: unknown[] = [];
  let error: string | undefined;
  const rewritten = content.replace(SHORTHAND_AT, (tag, userId: string, name: string) => {
    if (!AAD_GUID.test(userId)) {
      error ??= `Mention user-id "${userId}" is not an Entra object id (GUID). Use the userId from list-team-members or list-chat-members.`;
      return tag;
    }
    const id = nextId++;
    added.push({
      id,
      mentionText: name,
      mentioned: { user: { id: userId, displayName: name, userIdentityType: 'aadUser' } },
    });
    return `<at id="${id}">${name}</at>`;
  });

  if (error) return { body, error };
  if (added.length === 0) return { body };
  return {
    body: { ...body, body: { ...itemBody, content: rewritten }, mentions: [...existing, ...added] },
  };
}
```

- [ ] **Step 4: Run to verify pass**

Run: `npx vitest run src/__tests__/graph-tools.test.ts -t rewriteMentionShorthand`
Expected: PASS (9 tests).

- [ ] **Step 5: Commit**

```bash
git add src/graph-tools.ts src/__tests__/graph-tools.test.ts
git commit -m "feat(tools): add Teams mention shorthand rewriter"
```

### Task 4: Wire the rewriter into the seven message tools

**Files:**

- Modify: `src/graph-tools.ts` (one site)
- Modify: `src/endpoints.json` (seven tips)
- Test: `src/__tests__/graph-tools.test.ts`

**Interfaces:**

- Consumes: `MENTION_SHORTHAND_TOOLS`, `rewriteMentionShorthand` from Task 3.

- [ ] **Step 1: Write the failing tests**

```ts
describe('mention shorthand (tool wiring)', () => {
  const U1 = '11111111-2222-3333-4444-555555555555';
  const aliases = [
    ['send-channel-message', 'post'],
    ['reply-to-channel-message', 'post'],
    ['update-channel-message', 'patch'],
    ['update-channel-message-reply', 'patch'],
    ['send-chat-message', 'post'],
    ['reply-to-chat-message', 'post'],
    ['update-chat-message', 'patch'],
  ] as const;

  async function call(alias: string, method: string, body: unknown) {
    mockEndpoints.push(
      makeEndpoint({
        alias,
        method,
        path: '/chats/{chat-id}/messages',
        parameters: [{ name: 'body', type: 'Body', schema: z.object({}).passthrough() }],
      })
    );
    mockEndpointsJson = [
      makeConfig({ toolName: alias, method, pathPattern: '/chats/{chat-id}/messages' }),
    ];
    const graphClient = createMockGraphClient([
      { content: [{ type: 'text', text: JSON.stringify({ id: 'm1' }) }] },
    ]);
    const server = createMockServer();
    const { registerGraphTools } = await loadModule();
    registerGraphTools(server as any, graphClient as any);
    const result = await server.tools.get(alias)!.handler({ 'chat-id': 'c1', body, confirm: true });
    return { result, graphClient };
  }
  const sent = (graphClient: any) => JSON.parse(graphClient.graphRequest.mock.calls[0][1].body);

  it.each(aliases)('rewrites the shorthand on %s', async (alias, method) => {
    const { graphClient } = await call(alias, method, {
      body: { contentType: 'html', content: `Hi <at user-id="${U1}">Jane</at>` },
    });
    const out = sent(graphClient);
    expect(out.body.content).toBe('Hi <at id="0">Jane</at>');
    expect(out.mentions).toEqual([
      {
        id: 0,
        mentionText: 'Jane',
        mentioned: { user: { id: U1, displayName: 'Jane', userIdentityType: 'aadUser' } },
      },
    ]);
  });

  it('returns a tool error and sends nothing for a bad user-id', async () => {
    const { result, graphClient } = await call('send-chat-message', 'post', {
      body: { contentType: 'html', content: '<at user-id="jane">Jane</at>' },
    });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('not an Entra object id');
    expect(graphClient.graphRequest).not.toHaveBeenCalled();
  });

  it('does not rewrite on a tool outside the seven', async () => {
    const { graphClient } = await call('create-draft-email', 'post', {
      body: { contentType: 'html', content: `<at user-id="${U1}">Jane</at>` },
    });
    expect(sent(graphClient).body.content).toBe(`<at user-id="${U1}">Jane</at>`);
  });

  it('lists only tools that exist in endpoints.json and documents the shorthand', () => {
    const endpoints = JSON.parse(
      readFileSync(join(__dirname, '..', 'endpoints.json'), 'utf8')
    ) as Array<{ toolName: string; llmTip?: string }>;
    for (const [alias] of aliases) {
      const entry = endpoints.find((e) => e.toolName === alias);
      expect(entry, alias).toBeDefined();
      expect(entry!.llmTip, alias).toContain('user-id=');
    }
  });
});
```

If `makeEndpoint` rejects `method: 'patch'` in this harness, check how existing tests build PATCH endpoints (`grep -n "method: 'patch'" src/__tests__/graph-tools.test.ts`) and mirror them.

- [ ] **Step 2: Run to verify failure**

Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "mention shorthand"`
Expected: FAIL (body sent unchanged; tips lack `user-id=`).

- [ ] **Step 3: Implement**

In `src/graph-tools.ts`, directly after the `if (Object.keys(strayBodyFields).length > 0) { ... }` block and before the `// Defense-in-depth: the calendar delta tools` comment:

```ts
if (MENTION_SHORTHAND_TOOLS.has(tool.alias)) {
  const rewritten = rewriteMentionShorthand(body);
  if (rewritten.error) {
    return {
      content: [{ type: 'text', text: JSON.stringify({ error: rewritten.error }) }],
      isError: true,
    };
  }
  body = rewritten.body;
}
```

Update the seven tips with a one-off script (keeps the file's formatting via prettier). Save as `scratch/set-tips.mjs` outside the repo tree or in the scratchpad dir, run from the worktree root, do not commit it:

```js
import { readFileSync, writeFileSync } from 'node:fs';
const path = 'src/endpoints.json';
const endpoints = JSON.parse(readFileSync(path, 'utf8'));
const aliases = [
  'send-channel-message',
  'reply-to-channel-message',
  'update-channel-message',
  'update-channel-message-reply',
  'send-chat-message',
  'reply-to-chat-message',
  'update-chat-message',
];
const lead =
  'To @mention someone, put <at user-id="<aad-user-id>">Display Name</at> in the html content (contentType html); the server numbers the tag and builds the mentions array. Get the id from list-team-members or list-chat-members (name param). ';
for (const alias of aliases) {
  const entry = endpoints.find((e) => e.toolName === alias);
  if (!entry) throw new Error(`missing ${alias}`);
  entry.llmTip = lead + (entry.llmTip ?? 'Use contentType html in the body.');
}
writeFileSync(path, JSON.stringify(endpoints, null, 2) + '\n');
```

Run: `node <path-to>/set-tips.mjs && npx prettier --write src/endpoints.json && git diff --stat src/endpoints.json`
Expected: only the seven `llmTip` lines changed. If more lines moved, `git checkout src/endpoints.json` and edit the seven tips by hand instead.

The existing hand-built-body instructions stay in each tip after the new lead sentence, as the fallback.

- [ ] **Step 4: Run to verify pass, then the whole suite**

Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "mention shorthand"` then `npm test`
Expected: PASS, no regressions.

- [ ] **Step 5: Verify and commit**

Run: `npx prettier --write src/graph-tools.ts src/endpoints.json src/__tests__/graph-tools.test.ts && npm run verify`
Expected: all green.

```bash
git add -A
git commit -m "feat(tools): accept <at user-id> mention shorthand on Teams message tools"
```

- [ ] **Step 6: STOP and ask the user** before `git push` / opening the PR. PR text: what, why, tests. No AI attribution.

---

## Self-review

- **Spec coverage:** PR A name param, stripped before Graph, displayName+email substring, post-projection, post-merge, empty result, `llmTip` → Tasks 1-2. PR B seven tools, numbering above existing ids, mentions merge, pass-through, bad-GUID error, html only, `llmTip` → Tasks 3-4. Process: branch from `upstream/main`, full suite before PR, no attribution → Setup and Global Constraints.
- **Placeholders:** none. Two steps ask the implementer to check a harness detail (`createMockServer` schema exposure, PATCH endpoint construction) with the exact command to run.
- **Names:** `filterMembersByName`, `MEMBER_NAME_FILTER_TOOLS`, `MEMBER_NAME_PARAM_DESCRIPTION`, `rewriteMentionShorthand`, `MENTION_SHORTHAND_TOOLS` are used identically across tasks.
