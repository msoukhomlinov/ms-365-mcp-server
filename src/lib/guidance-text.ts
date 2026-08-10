/**
 * Keeping guidance text honest about which tools exist.
 *
 * Tool descriptions, llmTips and `initialize.instructions` are static strings
 * written once and emitted in every mode, but the registered tool set is not
 * static: `--attachment-proxy` deletes download-bytes, get-download-url and
 * get-mail-message-mime and puts read-document in their place. Every sentence
 * that says "call download-bytes with target=..." then survives into a server
 * that has no such tool, and a model following the guidance calls a name that
 * is not there.
 *
 * The fix has to hold for text nobody has written yet — the next llmTip naming
 * a byte tool is a matter of when, not if — so this drops the sentences that
 * name a departed tool rather than matching the sentences we happen to have
 * today. Dropping, not rewriting: a mechanical substitution of read-document
 * for download-bytes reads plausibly but lies about the resulting behaviour
 * ("returns base64", "with no Authorization header", "to retrieve the current
 * photo"), and confidently false guidance is worse than none. What the removed
 * sentences were for is restored once, by `replacement`.
 */

/**
 * Sentence boundary: a period followed by whitespace and the start of a new
 * sentence. Requiring an uppercase letter or an opening bracket after the
 * space keeps `@microsoft.graph.downloadUrl`, `ProfilePhoto.ReadWrite.All`,
 * `/$value.` and `{id}.` intact — those periods are either unspaced or
 * followed by lowercase.
 */
const SENTENCE_BOUNDARY = /(?<=\.)\s+(?=[A-Z(])/;

/**
 * A tool name as it appears in prose. Matched on word boundaries, so
 * `download-bytes` does not match inside `download-bytes-to-file`, and
 * separator-agnostically, because MCP clients render these names with
 * underscores (`download_bytes`) and guidance is sometimes written that way.
 */
function toolNamePattern(name: string): RegExp {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/-/g, '[-_]');
  // Case-insensitive: a sentence that opens with a tool name gets capitalised
  // when guidance is assembled, and a case-sensitive match would let
  // "Download-bytes ..." slip past the very check that exists to catch it.
  // Safe here because these names are hyphenated and do not occur as prose.
  return new RegExp(`(?<![A-Za-z0-9_-])${escaped}(?![A-Za-z0-9_-])`, 'i');
}

/** Whether `text` names any of `toolNames`. */
export function mentionsToolName(text: string, toolNames: readonly string[]): boolean {
  return toolNames.some((name) => toolNamePattern(name).test(text));
}

/**
 * Guidance to offer in place of what was dropped.
 *
 * `whenDropped` keeps the offer honest: the replacement is appended only if a
 * sentence naming one of those tools is what went. A sentence dropped for an
 * unrelated reason — a preset that excluded a cross-referenced tool, say — gets
 * no substitute, because the substitute would answer a question nobody asked.
 */
export interface GuidanceReplacement {
  text: string;
  whenDropped: readonly string[];
}

/**
 * `text` with every sentence naming a tool in `suppressed` removed, and
 * `replacement` appended once (and only once, however many sentences went) when
 * its `whenDropped` condition is met. Returns `text` unchanged when `suppressed`
 * is empty or nothing matches, which is the common case — no configuration pays
 * for this but the one that needs it.
 */
export function stripStaleToolGuidance(
  text: string,
  suppressed: readonly string[],
  replacement?: GuidanceReplacement
): string {
  if (!text || suppressed.length === 0) return text;
  if (!mentionsToolName(text, suppressed)) return text;

  const kept: string[] = [];
  const dropped: string[] = [];
  for (const sentence of text.split(SENTENCE_BOUNDARY)) {
    (mentionsToolName(sentence, suppressed) ? dropped : kept).push(sentence);
  }

  const parts = kept.map((s) => s.trim()).filter((s) => s.length > 0);
  if (
    replacement &&
    mentionsToolName(dropped.join(' '), replacement.whenDropped) &&
    !parts.includes(replacement.text.trim())
  ) {
    parts.push(replacement.text.trim());
  }
  return parts.join(' ');
}
