/**
 * Process-wide handle on the configured document-conversion proxy.
 *
 * The same module-singleton shape as `attachment-minting.ts`, deliberately and
 * for the same reason: exactly one tool reads it, and `registerGraphTools` and
 * `registerDiscoveryTools` already take eleven and twelve positional arguments.
 * Threading a client through both, plus `UtilityToolContext`, plus every call
 * site and every test that builds one, spreads a single-consumer dependency
 * across a dozen signatures.
 *
 * The URL is carried alongside the client rather than read back out of it. The
 * client's interface does not expose it, and the `proxy_unreachable` warn log
 * has to name the address that failed -- "the proxy is down" in `docker logs`
 * without saying which address was dialled is barely better than silence.
 *
 * Configured once, from `server.ts`, before any tool can run. Null is the state
 * every stdio run and every HTTP run without `--attachment-proxy` stays in.
 */

import type { AttachmentProxyClient } from './attachment-proxy.js';

export interface AttachmentProxyRuntime {
  client: AttachmentProxyClient;
  /** The URL the client was built for. Diagnostics only; never sent anywhere. */
  url: string;
}

/**
 * 60 seconds. Sized against measured behaviour rather than a round number: a
 * 195 KB PDF converted in 0.112 s live, so this is ~500x the observed cost and
 * exists to bound a wedged proxy rather than to be reached by a slow one.
 */
export const ATTACHMENT_PROXY_TIMEOUT_MS = 60_000;

let current: AttachmentProxyRuntime | null = null;

export function configureAttachmentProxy(runtime: AttachmentProxyRuntime | null): void {
  current = runtime;
}

export function getAttachmentProxy(): AttachmentProxyRuntime | null {
  return current;
}

/** Test helper -- drops any configured state so cases cannot leak into each other. */
export function resetAttachmentProxy(): void {
  current = null;
}
