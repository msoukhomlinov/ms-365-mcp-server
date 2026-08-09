/**
 * Short-TTL, budgeted tickets for Graph byte resources that expose no
 * pre-authenticated download URL of their own.
 *
 * A ticket is a capability and nothing else: it names one Graph path and one
 * account, it may be redeemed a small fixed number of times, and it expires. It
 * carries no credential -- the redemption route authenticates to Graph with the
 * server's own token, the same way every tool does. Holding a ticket therefore
 * authorises at most `MAX_REDEMPTIONS` authenticated GETs of exactly one
 * resource inside one short TTL.
 *
 * **Why a budget and not a single use.** Single-use was the original grant, and
 * it was too small to express the flow the feature exists for. A redemption was
 * spent the instant it *started*, so a Graph 5xx, an uninitialised client, or a
 * mid-stream abort left the caller holding a dead URL and no bytes -- measured
 * live as a Graph 404 on `/me/photo/$value` that answered 502 and burnt the
 * ticket. It also made the document sidecar's own documented flow impossible:
 * probing a document and then converting it are two fetches of one URL, as is
 * every pagination continuation, so a probe always killed the convert that was
 * supposed to follow it.
 *
 * A budget keeps what single-use was actually protecting. The threat is a URL
 * that leaks -- into a transcript, a log, an error message -- being replayed by
 * whoever finds it. What bounds that is the *product* of a short TTL and a small
 * finite count, not the count being exactly one: three fetches inside two
 * minutes is the same order of exposure as one, and both are unrelated to the
 * unbounded replay that dropping the limit entirely would allow.
 *
 * Every redemption costs one, whether or not it produced bytes. Refunding
 * failures was considered and rejected twice over: it would make the advertised
 * count untrue ("three, unless the failures were the right kind"), which is the
 * worst property an interface a model reads can have, and a target that fails
 * every time would then be re-fetchable without limit for the whole TTL.
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

export interface AttachmentTicket {
  /** Relative Graph path, exactly as the minting tool validated it. */
  readonly target: string;
  /** Account this ticket was minted for; undefined in single-account mode. */
  readonly accountName: string | undefined;
  /** Epoch milliseconds after which this ticket is dead. */
  readonly expiresAtMs: number;
  /**
   * Content-Type the minting tool learned from Graph's own metadata for this
   * resource, at mint time -- independent of whatever Content-Type header the
   * `/$value` fetch itself carries at redemption time. `null` when the target
   * has no metadata resource to probe (drive/SharePoint content, meeting
   * recordings, an unmintable target) or the probe found nothing specific.
   *
   * The redemption route (`attachment-route.ts`) is what actually chooses
   * between this and the stream's own header -- see its `isGenericContentType`
   * precedence -- so this field is deliberately just data, not a decision:
   * a stale or wrong probe can never cost more than falling back to today's
   * behaviour.
   */
  readonly probedContentType: string | null;
  /** Graph's own `name` for this resource at mint time, same caveats as above. */
  readonly probedName: string | null;
}

/** Mutable store entry. `AttachmentTicket` is the part a redeemer may see. */
interface StoredTicket extends AttachmentTicket {
  /** Redemptions not yet handed out. Reaches 0 only as the entry is deleted. */
  remaining: number;
  /** A redemption is currently holding this ticket; see `redeem`. */
  inFlight: boolean;
}

/**
 * An exclusive hold on one redemption of one ticket.
 *
 * The hold lasts until `release()`, which the redeeming route must call on
 * every exit path -- a lease that is never released pins the ticket for the
 * rest of its TTL, which fails closed (later fetches 404) but wastes the
 * caller's remaining budget.
 */
export interface AttachmentTicketLease extends AttachmentTicket {
  /**
   * Redemptions left *after* this one. Diagnostics only: it must never reach a
   * response body, because "2 left" distinguishes a real ticket from a guessed
   * one and that is exactly the oracle the uniform 404 exists to deny.
   */
  readonly remaining: number;
  /** Drop the exclusive hold. Idempotent, and safe after the ticket is gone. */
  release(): void;
}

/**
 * Redemptions one minted URL is worth.
 *
 * Three, sized to the flow it has to survive rather than to a round number: the
 * document sidecar's own advice is to probe a large document and then convert
 * it (two fetches of one URL), and one of those two is allowed to fail and be
 * retried. A fourth would buy a second pagination continuation, which the
 * 120-second TTL usually will not reach anyway; the TTL, not this number, is
 * what ends a long read.
 */
export const MAX_REDEMPTIONS = 3;

/**
 * Cap on live tickets. A ticket is ~200 bytes, so this bounds the store at a
 * few hundred KB -- but the reason for a cap is not memory, it is that an agent
 * in a retry loop should hit a refusal it can report rather than grow the
 * process without limit. Minting refuses when full, after sweeping; it never
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

  mint(
    target: string,
    accountName: string | undefined,
    nowMs: number = Date.now(),
    probe?: { contentType?: string | null; name?: string | null }
  ): { id: string; expiresAtMs: number } {
    this.sweep(nowMs);
    if (this.tickets.size >= MAX_LIVE_TICKETS) {
      throw new TicketStoreFullError(MAX_LIVE_TICKETS);
    }
    const id = randomBytes(TICKET_BYTES).toString('base64url');
    const expiresAtMs = nowMs + this.ttlSeconds * 1000;
    this.tickets.set(id, {
      target,
      accountName,
      expiresAtMs,
      remaining: MAX_REDEMPTIONS,
      inFlight: false,
      probedContentType: probe?.contentType ?? null,
      probedName: probe?.name ?? null,
    });
    return { id, expiresAtMs };
  }

  /**
   * Take one redemption of a ticket, or return undefined.
   *
   * One `undefined` for every failure -- unknown id, exhausted, expired, and
   * "another request is redeeming it right now". The caller answers 404 to all
   * four, so a probe cannot use the response to tell "never existed" from
   * "real, but spent", which would confirm a guessed id. The in-flight case
   * matters most here: it is the only refusal an attacker could *provoke*
   * (by racing a legitimate fetch), and it has to look like the rest.
   *
   * **This is how two simultaneous requests are stopped from both streaming.**
   * The whole method is synchronous -- no `await`, so no interleaving point --
   * which makes test-and-set atomic with respect to every other request in the
   * process. The second request sees `inFlight` and is refused. It stays
   * refused until `release()`, i.e. until the first request's stream has
   * finished, so at most one response is ever in flight per ticket even though
   * the ticket outlives its first use.
   *
   * The budget is spent here rather than on success, and the last one deletes
   * the entry as it is handed out, so the final redemption behaves exactly as
   * single-use did: an exception anywhere on the streaming path cannot leave a
   * redemption that was already counted available again.
   */
  redeem(id: string, nowMs: number = Date.now()): AttachmentTicketLease | undefined {
    this.sweep(nowMs);
    const ticket = this.tickets.get(id);
    if (!ticket) return undefined;
    // No second expiry check here: `sweep` above ran against this same `nowMs`
    // and already removed anything at or past its expiry, so a surviving entry
    // is live by construction -- and because expiry is enforced by that sweep
    // rather than by the budget, an unspent ticket past its TTL is still dead.
    // A re-check would be unreachable code asserting a guarantee the sweep
    // already provides, and with one captured timestamp there is no race.
    if (ticket.inFlight) return undefined;

    ticket.inFlight = true;
    ticket.remaining -= 1;
    // `remaining` is >= 1 on entry, because the entry is removed at the moment
    // its last redemption is handed out rather than left at zero to be refused
    // later. That keeps "present in the map" and "redeemable" the same fact.
    if (ticket.remaining <= 0) this.tickets.delete(id);

    let released = false;
    return {
      target: ticket.target,
      accountName: ticket.accountName,
      expiresAtMs: ticket.expiresAtMs,
      remaining: ticket.remaining,
      probedContentType: ticket.probedContentType,
      probedName: ticket.probedName,
      release: () => {
        if (released) return;
        released = true;
        // Harmless when the entry was already deleted (last redemption) or
        // swept (expired mid-stream): it mutates an object nothing can reach.
        ticket.inFlight = false;
      },
    };
  }

  /** Live ticket count, for tests and diagnostics. Never logged with ids. */
  size(nowMs: number = Date.now()): number {
    this.sweep(nowMs);
    return this.tickets.size;
  }

  /**
   * Redemptions this ticket has left, or 0 if it is unknown, exhausted or
   * expired -- the same deliberate conflation `redeem` makes.
   *
   * Diagnostics and tests only. Nothing serves this over the wire: an endpoint
   * answering "0 vs 2" for an id would be precisely the probe oracle the
   * uniform 404 denies.
   */
  redemptionsLeft(id: string, nowMs: number = Date.now()): number {
    this.sweep(nowMs);
    return this.tickets.get(id)?.remaining ?? 0;
  }

  clear(): void {
    this.tickets.clear();
  }
}
