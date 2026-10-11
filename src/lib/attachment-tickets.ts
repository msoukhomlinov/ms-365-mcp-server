/**
 * Short-TTL, single-redemption tickets for Graph byte resources that expose no
 * pre-authenticated download URL of their own.
 *
 * A ticket is a capability and nothing else: it names one Graph path and one
 * identity, it is redeemable once unless the minting caller asked for a small
 * finite count (see `maxRedemptions`), and it expires. The identity is an account
 * in this server's own token cache, or, when Graph identity came from the
 * request, the token that request read with. That token stays in the record
 * and never reaches the URL. Holding a ticket therefore authorises exactly one
 * authenticated GET of exactly one resource, which is the smallest grant that
 * makes an out-of-band fetch possible at all.
 *
 * Memory-only, deliberately. Persisting tickets would mean a redeemable
 * capability surviving a restart, and re-reading it from disk is a second place
 * for it to leak; a ticket outliving the process it was minted in has no
 * legitimate use when the TTL is measured in minutes.
 */

import { randomBytes } from 'node:crypto';
import { ATTACHMENT_ROUTE, type AttachmentUrlConfig } from './attachment-url-config.js';
import { signUrl } from './url-signing.js';

/** Query parameter carrying the ticket id. */
export const TICKET_PARAM = 't';

/**
 * Build the signed, redeemable URL for a minted ticket.
 *
 * **The ticket travels in the query, never the path**, and that is a hard
 * requirement of the verifying sidecar rather than a style choice: docglean's
 * error messages keep a fetched URL's path (so an operator can tell which
 * document failed) and strip its query. A ticket in the path would be signed
 * just as correctly and would also land in every one of those messages.
 */
export function buildAttachmentUrl(
  config: AttachmentUrlConfig,
  ticketId: string,
  nowMs: number = Date.now()
): string {
  const url = new URL(ATTACHMENT_ROUTE, config.base);
  url.searchParams.set(TICKET_PARAM, ticketId);
  return signUrl(
    url.toString(),
    { key: config.key, keyId: config.keyId, ttlSeconds: config.ttlSeconds },
    nowMs
  );
}

/** Any origin and version prefix will do; only whether parsing changes the path matters. */
const PROBE_ORIGIN = 'https://graph.invalid';
const PROBE_PREFIX = '/v1.0';

/**
 * Whether a minted target survives URL parsing unchanged.
 *
 * The mint-site checks match on a suffix, which says nothing about what precedes it, and
 * `performRequest` later concatenates the target onto the Graph origin and lets WHATWG
 * parse the result. Anything the parser rewrites on the way makes the path that was
 * validated and the path that gets fetched two different things.
 *
 * This asks that question directly rather than enumerating the ways it can happen, because
 * enumerating them kept coming up short: dot segments resolve away, a fragment never
 * reaches the wire, and TAB, LF and CR are *deleted before* dot segments resolve, so a `..`
 * split by one of them survives a segment comparison and is reassembled by the parser.
 * Round-tripping catches that class whole, including the encoded and control-character
 * spellings, and stays correct if the parser grows another normalisation.
 *
 * `//` is rejected separately. With the version prefix in front it stays an ordinary path
 * segment, so the round trip accepts it, but it is only inert for as long as the caller
 * keeps prefixing something -- without that it parses as an authority and points off-origin.
 */
export function isPlainGraphPath(target: string): boolean {
  if (!target.startsWith('/')) return false;
  if (target.includes('//')) return false;
  let resolved: URL;
  try {
    resolved = new URL(PROBE_ORIGIN + PROBE_PREFIX + target);
  } catch {
    return false;
  }
  return (
    resolved.origin === PROBE_ORIGIN &&
    resolved.search === '' &&
    resolved.hash === '' &&
    resolved.pathname === PROBE_PREFIX + target
  );
}

/**
 * The same round trip for a path filled in from a tool's own template. Compared decoded,
 * because skipEncoding parameters reach the path raw and the parser percent-encodes a
 * space or a non-ASCII letter in one without sending the request anywhere else.
 */
export function isUnalteredGraphPath(path: string): boolean {
  // A "%" that starts no escape is literal text ("100%"), which the parser leaves alone
  const decode = (s: string) => decodeURIComponent(s.replace(/%(?![0-9a-f]{2})/gi, '%25'));
  try {
    const resolved = new URL(PROBE_ORIGIN + PROBE_PREFIX + path);
    return (
      resolved.origin === PROBE_ORIGIN &&
      resolved.search === '' &&
      resolved.hash === '' &&
      decode(resolved.pathname) === decode(PROBE_PREFIX + path)
    );
  } catch {
    return false;
  }
}

/**
 * Who the bytes are fetched as. Two kinds rather than two optional fields, so
 * redemption has to pick a branch and a request-token ticket cannot reach the
 * token cache by way of a missing value.
 */
export type TicketIdentity =
  /** Account in this server's token cache; undefined in single-account mode. */
  | { readonly kind: 'server-account'; readonly accountName: string | undefined }
  /** Graph token of the request that minted the ticket. */
  | { readonly kind: 'request-token'; readonly accessToken: string };

/**
 * What Graph's own metadata said about a mail or event attachment at mint time. Plain data:
 * the redemption route decides whether to use it, so a stale or wrong probe can never cost
 * more than falling back to what the fetch itself reports.
 */
export type DownloadProbe = {
  readonly probedContentType?: string;
  readonly probedName?: string;
};

/**
 * What redeeming the ticket does. A download ticket authorises one GET of
 * `target`; an upload ticket authorises one PUT of a file that becomes an
 * attachment in the `target` collection (a draft message's or an event's
 * `/attachments`). Two kinds rather than a flag so a redemption handler has to
 * pick the branch it was written for and cannot serve the other by accident.
 */
export type TicketPurpose =
  | ({ readonly purpose: 'download' } & DownloadProbe)
  | {
      readonly purpose: 'upload';
      /** Attachment name as Graph will store it. */
      readonly name: string;
      /** MIME type recorded on the attachment. */
      readonly contentType: string;
    };

export type AttachmentTicket = TicketIdentity &
  TicketPurpose & {
    /** Relative Graph path, exactly as the minting tool validated it. */
    readonly target: string;
    /** Epoch milliseconds after which this ticket is dead. */
    readonly expiresAtMs: number;
  };

/**
 * Most redemptions a single download ticket can be minted for. Sized to the flow that
 * needs more than one: probe a document, convert it, retry one failure.
 */
export const MAX_REDEMPTIONS = 3;

/** Mutable store entry. `AttachmentTicket` is the part a redeemer may see. */
type StoredTicket = AttachmentTicket & {
  /** Redemptions not yet handed out. Reaches 0 only as the entry is deleted. */
  remaining: number;
  /** A redemption is holding this ticket; see `redeem`. */
  inFlight: boolean;
};

/**
 * An exclusive hold on one redemption of one ticket. The redeeming route must call
 * `release()` on every exit path: a lease that is never released pins a multi-use ticket
 * for the rest of its TTL, which fails closed (later fetches 404) but wastes its budget.
 */
export type AttachmentTicketLease = AttachmentTicket & {
  /** Drop the exclusive hold. Idempotent, and safe after the ticket is gone. */
  release(): void;
};

/** Upload metadata a minting tool supplies; the store adds nothing to it. */
export type UploadSpec = { readonly name: string; readonly contentType: string };

/**
 * Cap on live tickets. A ticket is a few KB at most, the token being nearly all
 * of it, so this bounds the store at about a MB -- but the reason for a cap is
 * not memory, it is that an agent in a retry loop should hit a refusal it can
 * report rather than grow the process without limit. Minting refuses when full,
 * after sweeping; it never
 * evicts a live ticket, because evicting the oldest would let a caller minting
 * in a loop invalidate tickets someone else is about to redeem.
 */
const MAX_LIVE_TICKETS = 256;

/** 32 bytes of CSPRNG output -- the ticket id is the whole capability. */
const TICKET_BYTES = 32;

export class TicketStoreFullError extends Error {
  constructor(public readonly limit: number) {
    super(`No ticket slots available (limit ${limit}); retry once outstanding tickets expire.`);
    this.name = 'TicketStoreFullError';
  }
}

export class AttachmentTicketStore {
  private readonly tickets = new Map<string, StoredTicket>();

  constructor(private readonly ttlSeconds: number) {}

  /** Drop every expired ticket. Called before each mint and each redemption. */
  private sweep(nowMs: number): void {
    for (const [id, ticket] of this.tickets) {
      if (ticket.expiresAtMs <= nowMs) this.tickets.delete(id);
    }
  }

  private add(
    target: string,
    identity: TicketIdentity,
    nowMs: number,
    purpose: TicketPurpose = { purpose: 'download' },
    maxRedemptions = 1
  ): { id: string; expiresAtMs: number } {
    if (
      !Number.isInteger(maxRedemptions) ||
      maxRedemptions < 1 ||
      maxRedemptions > MAX_REDEMPTIONS
    ) {
      throw new RangeError(`maxRedemptions must be an integer from 1 to ${MAX_REDEMPTIONS}`);
    }
    this.sweep(nowMs);
    if (this.tickets.size >= MAX_LIVE_TICKETS) {
      throw new TicketStoreFullError(MAX_LIVE_TICKETS);
    }
    const id = randomBytes(TICKET_BYTES).toString('base64url');
    const expiresAtMs = nowMs + this.ttlSeconds * 1000;
    this.tickets.set(id, {
      ...identity,
      ...purpose,
      target,
      expiresAtMs,
      remaining: maxRedemptions,
      inFlight: false,
    });
    return { id, expiresAtMs };
  }

  /**
   * Mint a ticket redeemed with this server's own token for `accountName`.
   * `maxRedemptions` (1 to MAX_REDEMPTIONS, default 1) is how many attempts it is good for;
   * `probe` is what Graph's metadata said about the target, if the minting caller asked.
   */
  mint(
    target: string,
    accountName: string | undefined,
    nowMs: number = Date.now(),
    maxRedemptions = 1,
    probe: DownloadProbe = {}
  ): { id: string; expiresAtMs: number } {
    return this.add(
      target,
      { kind: 'server-account', accountName },
      nowMs,
      { purpose: 'download', ...probe },
      maxRedemptions
    );
  }

  /** Mint a ticket redeemed with `accessToken` and nothing else; `maxRedemptions` and `probe` as for `mint`. */
  mintWithToken(
    target: string,
    accessToken: string,
    nowMs: number = Date.now(),
    maxRedemptions = 1,
    probe: DownloadProbe = {}
  ): { id: string; expiresAtMs: number } {
    if (!accessToken) throw new Error('A request-token ticket needs a token');
    return this.add(
      target,
      { kind: 'request-token', accessToken },
      nowMs,
      { purpose: 'download', ...probe },
      maxRedemptions
    );
  }

  /** Mint an upload ticket attached as this server's own token for `accountName`. */
  mintUpload(
    target: string,
    accountName: string | undefined,
    upload: UploadSpec,
    nowMs: number = Date.now()
  ): { id: string; expiresAtMs: number } {
    return this.add(target, { kind: 'server-account', accountName }, nowMs, {
      purpose: 'upload',
      ...upload,
    });
  }

  /** Mint an upload ticket attached with `accessToken` and nothing else. */
  mintUploadWithToken(
    target: string,
    accessToken: string,
    upload: UploadSpec,
    nowMs: number = Date.now()
  ): { id: string; expiresAtMs: number } {
    if (!accessToken) throw new Error('A request-token ticket needs a token');
    return this.add(target, { kind: 'request-token', accessToken }, nowMs, {
      purpose: 'upload',
      ...upload,
    });
  }

  /**
   * Take one redemption of a ticket, or return undefined.
   *
   * One `undefined` for every failure -- unknown id, exhausted, expired, and "another
   * request is redeeming it right now". The caller answers 404 to all of them, so a probe
   * cannot use the response to tell "never existed" from "real, but spent". The in-flight
   * case matters most: it is the only refusal an attacker could provoke, by racing a
   * legitimate fetch, and it has to look like the rest.
   *
   * Every redemption costs one whether or not it produced bytes, so a target that fails
   * every time cannot be re-fetched without limit. The budget is spent here, and the last
   * one deletes the entry as it is handed out: a single-use ticket is burnt before the
   * value is returned, exactly as it always was, and an exception on the streaming path
   * cannot leave a counted redemption available again.
   *
   * The method is synchronous, so test-and-set on `inFlight` is atomic against every other
   * request in the process: at most one response is in flight per ticket. A request
   * refused for that reason spends nothing.
   */
  redeem(id: string, nowMs: number = Date.now()): AttachmentTicketLease | undefined {
    this.sweep(nowMs);
    const stored = this.tickets.get(id);
    if (!stored || stored.inFlight) return undefined;

    stored.inFlight = true;
    stored.remaining -= 1;
    if (stored.remaining <= 0) this.tickets.delete(id);

    const { remaining: _remaining, inFlight: _inFlight, ...ticket } = stored;
    let released = false;
    return {
      ...ticket,
      // Once only: a repeat after a later lease took over must not drop that lease's hold.
      // Harmless once the entry is deleted or swept: it mutates an object nothing reaches.
      release: () => {
        if (released) return;
        released = true;
        stored.inFlight = false;
      },
    };
  }

  /** Live ticket count, for tests and diagnostics. Never logged with ids. */
  size(nowMs: number = Date.now()): number {
    this.sweep(nowMs);
    return this.tickets.size;
  }

  clear(): void {
    this.tickets.clear();
  }
}
