/**
 * Redemption route for server-minted attachment URLs.
 *
 * The URL a document-conversion sidecar fetches looks like:
 *
 *     GET /attachment?t=<ticket>&dgk=<key-id>&dgx=<expiry>&dgs=<signature>
 *
 * **This route ignores `dgk`/`dgx`/`dgs` entirely, and that is correct.** Those
 * three exist for the sidecar, which verifies them before it will dial a
 * private address at all; they are the sidecar's authorisation to *dial*, not
 * anyone's authorisation to *redeem*. What authorises redemption here is `t` --
 * a short-TTL capability this server minted and remembers, good for a small
 * fixed number of fetches (`MAX_REDEMPTIONS`) and no more. Checking
 * the signature here as well would buy nothing (the key is ours, so a valid
 * signature says only that we minted the URL, which the ticket already proves)
 * and would cost something real: it would couple redemption to the sidecar's
 * clock and to the key surviving a restart, turning two independent failures
 * into one.
 *
 * No Authorization header is required or read. The fetcher holds no Microsoft
 * credential -- that is the entire point of handing it a URL instead of bytes --
 * so the ticket is the only credential in play, and the response is streamed
 * with this server's own Graph token.
 */

import type { Handler, Request, Response } from 'express';
import { Readable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import logger from './logger.js';
import type GraphClient from './graph-client.js';
import type AuthManager from './auth.js';
import { type AttachmentTicketStore, TICKET_PARAM } from './lib/attachment-tickets.js';

export interface AttachmentRouteDeps {
  store: AttachmentTicketStore;
  getGraphClient: () => GraphClient | null;
  authManager: AuthManager;
}

/**
 * One body for every refusal.
 *
 * Unknown id, exhausted, expired, malformed, and "already being redeemed right
 * now" all answer an identical 404. A distinguishable response would confirm a
 * guessed ticket id -- "this one existed but is spent" is most of the way to
 * knowing an id is real -- and ticket ids are the whole capability.
 */
const NOT_FOUND_BODY = 'Not found';

function refuse(res: Response): void {
  res.status(404).type('text/plain').send(NOT_FOUND_BODY);
}

/**
 * Content-Type values that name no real format.
 *
 * `application/octet-stream` is `graph-client.ts`'s own fallback when Graph's
 * `/$value` response carries no `content-type` header at all -- so by the
 * time it reaches here it is indistinguishable from "Graph genuinely said
 * this", and either way a generic subtype matches nothing in a converter's
 * format table. `application/binary` is the same shape from a different
 * source (some older Exchange/Outlook clients write it into an attachment's
 * own stored MIME type when they had nothing better). The empty string is
 * defensive -- `downloadStream` should never hand back `''`, since it already
 * ORs against the same fallback, but treating it as generic rather than
 * "specific and empty" costs nothing and closes off a future refactor
 * removing that OR from silently reintroducing an unroutable header.
 * Parameters (`; charset=...`) are stripped before comparing, the same way
 * `isBinaryContentType` in graph-client.ts does it.
 */
const GENERIC_CONTENT_TYPES = new Set(['application/octet-stream', 'application/binary', '']);

function isGenericContentType(contentType: string): boolean {
  return GENERIC_CONTENT_TYPES.has(contentType.split(';')[0].trim().toLowerCase());
}

/**
 * Every character Node's own `http.OutgoingMessage.setHeader` accepts in a
 * header VALUE: horizontal tab, the printable ASCII range, and the raw Latin-1
 * supplement (0x80-0xFF) -- HTTP header values are historically ISO-8859-1,
 * not UTF-8, so Node does not reject bytes in that range even though they are
 * not ASCII. Anything else, including every code point past U+00FF, makes
 * `setHeader` throw `ERR_INVALID_CHAR` synchronously (verified directly
 * against a real `http.OutgoingMessage` in this file's test suite). A field
 * this route reads from Graph's OWN metadata JSON -- `name`, `contentType` --
 * is not constrained to this set at the source the way `stream.contentType`
 * and `stream.contentDisposition` are (those already passed through an actual
 * HTTP response header when Graph sent them, which could not have carried an
 * invalid byte in the first place), so anything probed has to be checked
 * before it reaches `setHeader`.
 */
const HEADER_SAFE_CHAR = /^[\t\x20-\x7e\x80-\xff]*$/;

/**
 * True when `contentType` can be set as a header value without Node throwing.
 * There is no RFC 5987-style fallback for Content-Type the way there is for a
 * filename: a MIME type is defined to be an ASCII token (RFC 6838), so a
 * probed value outside the header-safe range is not a real content-type
 * Graph would have produced -- it is unusable metadata, and the caller falls
 * back to the stream's own type (which is always header-safe already) rather
 * than transliterating something that was never going to be a valid MIME
 * type regardless.
 */
function isHeaderSafeContentType(contentType: string): boolean {
  return HEADER_SAFE_CHAR.test(contentType);
}

/**
 * Strip the characters that would let a name escape a quoted-string
 * parameter (`"`, and its escape character `\`) or, combined with anything
 * outside `HEADER_SAFE_CHAR`, inject a second header via CR/LF. Used both for
 * the ASCII-safe `filename=` fallback and, implicitly, made irrelevant for
 * `filename*=` -- `encodeURIComponent` percent-encodes `"` and `\` on its own,
 * so the extended form never needs this step.
 */
function quotedStringSafe(value: string): string {
  return value.replace(/[\r\n"\\]/g, '');
}

/**
 * An ASCII/Latin-1-safe stand-in for `name`, for the mandatory `filename=`
 * parameter RFC 6266 requires alongside any `filename*=` extension (a client
 * that does not understand the extended form falls back to this one, so it
 * has to be usable on its own, not empty).
 *
 * Drops every character `HEADER_SAFE_CHAR` would reject, plus the
 * quoted-string-breaking ones `quotedStringSafe` handles. When nothing
 * printable survives -- a name that is entirely outside Latin-1, e.g. an
 * all-CJK filename -- recovers at least the extension if the original name
 * has an ASCII one, which is exactly the signal docglean's own
 * extension-based format resolver would use; falls back to the bare
 * `attachment` disposition's own filename convention otherwise.
 */
function asciiSafeFilenameFallback(name: string): string {
  const safe = quotedStringSafe(name)
    .replace(/[^\t\x20-\x7e\x80-\xff]/g, '')
    .trim();
  const extension = /(\.[A-Za-z0-9]{1,10})$/.exec(name)?.[1] ?? '';
  // If what survived is nothing, or is nothing MORE than the extension itself
  // (every character of the actual name was outside the safe range), a bare
  // extension like `.pdf` is a worse fallback than a named one: it reads as a
  // hidden file with no name, not as "a PDF this server could not label".
  if (safe && safe !== extension) return safe;
  return extension ? `attachment${extension}` : safe || 'attachment';
}

/**
 * Percent-encode `str` per RFC 5987's `attr-char` grammar for use in an
 * `ext-value` (the `filename*=UTF-8''<this>` form).
 *
 * `encodeURIComponent` alone is not sufficient: it leaves `' ( ) *` unescaped
 * because they are "unreserved" for a URI component, but none of the four is
 * in `attr-char` (`ALPHA / DIGIT / "!" / "#" / "$" / "&" / "+" / "-" / "." /
 * "^" / "_" / "`" / "|" / "~"`), so a name containing any of them would
 * produce a value that fails RFC 5987's grammar even though it looks
 * plausible. `"` and `\` ARE covered by plain `encodeURIComponent` already
 * (neither is in its unreserved set), so `filename*=` needs no separate
 * quoted-string handling the way `filename=` does.
 */
function encodeRfc5987ValueChars(str: string): string {
  return encodeURIComponent(str).replace(
    /['()*]/g,
    (char) => '%' + char.charCodeAt(0).toString(16).toUpperCase()
  );
}

/**
 * A `Content-Disposition` header naming `name`, or the bare `attachment`
 * fallback when there is no name to give. `name` is untrusted -- it is
 * Graph's own attachment metadata, not this server's choice -- and, unlike
 * `stream.contentDisposition`, never passed through an actual HTTP header at
 * its source, so it can contain anything, including characters Node's own
 * `setHeader` rejects outright (see `HEADER_SAFE_CHAR`).
 *
 * Always emits an ASCII/Latin-1-safe `filename=` (RFC 6266 requires one
 * regardless, as the fallback for a client that ignores the extended form),
 * and additionally emits `filename*=UTF-8''<percent-encoded>` -- carrying the
 * real name -- whenever `name` contains anything outside that safe range.
 * Emitting only the extended form, or only a naively-encoded one, are both
 * real mistakes this function avoids on purpose: RFC 6266/5987 clients expect
 * both parameters together, and a bare `encodeURIComponent` swap would
 * produce a value invalid under RFC 5987's narrower `attr-char` set.
 */
function contentDispositionFor(name: string | null): string {
  if (!name) return 'attachment';
  const fallback = asciiSafeFilenameFallback(name);
  let value = `attachment; filename="${fallback}"`;
  if (!HEADER_SAFE_CHAR.test(name)) {
    value += `; filename*=UTF-8''${encodeRfc5987ValueChars(name)}`;
  }
  return value;
}

export function createAttachmentHandler(deps: AttachmentRouteDeps): Handler {
  return async (req: Request, res: Response): Promise<void> => {
    const raw = req.query[TICKET_PARAM];
    // Express parses a repeated `?t=a&t=b` into an array. Refuse rather than
    // picking one: two tickets in one request is not a shape any legitimate
    // caller produces, and silently taking the first would let an attacker
    // append a guess to a valid URL and learn from the timing which was used.
    if (typeof raw !== 'string' || raw.length === 0) {
      refuse(res);
      return;
    }

    // Takes one of the ticket's redemptions *and* an exclusive hold on it, in
    // one synchronous step. The hold is what keeps a ticket that survives its
    // first fetch from being streamed twice at once; it must be released on
    // every exit path below, hence the `finally`.
    const lease = deps.store.redeem(raw);
    if (!lease) {
      refuse(res);
      return;
    }

    try {
      const graphClient = deps.getGraphClient();
      if (!graphClient) {
        // The redemption is spent even though no byte was fetched, and that is
        // deliberate: a refund here would make the count this server advertises
        // untrue, and the caller still holds whatever budget is left. The tool
        // response tells the agent to retry the same URL, which is the right
        // move -- a fresh mint would meet the same uninitialised client.
        logger.error('Attachment redemption failed: Graph client is not initialised');
        res.status(503).type('text/plain').send('Service unavailable');
        return;
      }

      let stream: Awaited<ReturnType<GraphClient['downloadStream']>>;
      try {
        let accessToken: string | undefined;
        if (!deps.authManager.isOAuthModeEnabled()) {
          accessToken = await deps.authManager.getTokenForAccount(lease.accountName);
        }
        stream = await graphClient.downloadStream(lease.target, { accessToken });
      } catch (error) {
        // The target path is logged; the ticket id never is. The path is what an
        // operator needs to diagnose a failure and is not itself a capability --
        // reaching it still requires this server's Graph token.
        //
        // Unlike before, this no longer strands the caller: the ticket keeps its
        // remaining redemptions, so the same URL can simply be fetched again.
        logger.error(
          `Attachment redemption failed for ${lease.target}: ${(error as Error).message}`
        );
        res.status(502).type('text/plain').send('Upstream fetch failed');
        return;
      }

      res.status(200);
      // Precedence: a SPECIFIC Content-Type on THIS fetch's own response wins,
      // because it is what Graph is answering right now; the ticket's probed
      // type (learned from Graph's attachment metadata at mint time, see
      // `probeMailEventAttachment` in graph-tools.ts) is used only when the
      // stream itself carries nothing useful. Authoritative metadata beats a
      // generic default, but a specific response header beats a probe that
      // may be stale -- the probe ran once, at mint time; this fetch is
      // happening now and may be a retry against a target whose Graph-side
      // state has moved on.
      const probedContentType =
        lease.probedContentType && isHeaderSafeContentType(lease.probedContentType)
          ? lease.probedContentType
          : null;
      res.setHeader(
        'content-type',
        isGenericContentType(stream.contentType) && probedContentType
          ? probedContentType
          : stream.contentType
      );
      // Only declare a length that is a real, positive count of bytes. A
      // `content-length: 0` on a body we are about to stream is never correct
      // here: Node ends the response after zero bytes, the pipeline below then
      // rejects with "Premature close" too late to change the status, and the
      // peer reads a clean, well-formed, empty 200 that no retry logic will ever
      // question. Omitting the header instead lets Node chunk the body, which is
      // always safe. Anything not a positive integer -- null, 0, a negative, a
      // fraction -- is dropped rather than trusted.
      if (
        stream.contentLength !== null &&
        Number.isInteger(stream.contentLength) &&
        stream.contentLength > 0
      ) {
        res.setHeader('content-length', String(stream.contentLength));
      }
      // Graph's own filename when it gave one. Falling back to the ticket's
      // probed name (rather than the bare `attachment` default) hands
      // docglean a second, cheaper axis for free: its own format resolver
      // also reads an extension off Content-Disposition's filename, so a
      // real name can recover a correct route even when neither Content-Type
      // is specific. `attachment` either way, never `inline`: this endpoint
      // serves untrusted bytes from a mailbox, and a browser that wandered
      // onto the URL must not render an inline text/html attachment as a page
      // on this origin.
      res.setHeader(
        'content-disposition',
        stream.contentDisposition ?? contentDispositionFor(lease.probedName)
      );
      res.setHeader('cache-control', 'no-store');
      res.setHeader('x-content-type-options', 'nosniff');

      try {
        await pipeline(Readable.fromWeb(stream.body as never), res);
      } catch (error) {
        // Headers are already sent, so there is no status left to change. Destroy
        // rather than end, so the peer sees a truncated transfer instead of a
        // short body that looks complete.
        //
        // The redemption stays spent. An abort halfway is not a failure that
        // delivered nothing -- bytes left this server -- so treating it as one
        // and handing the redemption back would let a caller drain a mailbox
        // resource in unlimited partial reads off a single ticket.
        logger.error(`Attachment stream aborted for ${lease.target}: ${(error as Error).message}`);
        res.destroy();
      }
    } finally {
      // Runs on every path above, including a throw the route does not catch.
      // Until it does, this ticket refuses every other request with the same
      // uniform 404 an unknown id gets.
      lease.release();
    }
  };
}
