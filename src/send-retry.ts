import type { ILogger } from "@microsoft/teams.common";
import { describeError } from "./errors";

/**
 * Retry wrapper for outbound activity sends.
 *
 * Sends to smba.trafficmanager.net intermittently fail at the socket level
 * (ECONNRESET). By the time an answer is being sent it has already cost a
 * Retrieval API call and two or more model rounds, so dropping it on a transient
 * network blip is the most expensive possible failure in the turn.
 *
 * TRADEOFF: a send that succeeded server-side but whose response was lost will
 * be retried, and the user sees the message twice. That is deliberate --
 * a rare duplicate is much cheaper than a silently lost answer. If duplicates
 * ever become the bigger problem, the fix is idempotency keys on the send, not
 * removing the retry.
 */

// Initial attempt plus three retries. Beyond that the connection is not blipping,
// it is down, and further attempts just delay the error the caller needs to see.
const MAX_ATTEMPTS = 4;
const BASE_DELAY_MS = 250;
const MAX_DELAY_MS = 4000;

/**
 * Socket-level failures. These mean the request never got a reply, so there is
 * nothing about it that a retry would repeat incorrectly.
 */
const RETRYABLE_NETWORK_CODES = [
  "ECONNRESET",
  "ETIMEDOUT",
  "ECONNREFUSED",
  "ECONNABORTED",
  "EPIPE",
  "EAI_AGAIN",
  "ENOTFOUND",
  "ENETUNREACH",
  "EHOSTUNREACH",
  "ERR_STREAM_PREMATURE_CLOSE",
];

export function isRetryableSendError(err: unknown): boolean {
  const e = err as {
    code?: unknown;
    message?: unknown;
    statusCode?: unknown;
    cause?: { code?: unknown };
    response?: { status?: number };
  };

  const status =
    typeof e?.response?.status === "number"
      ? e.response.status
      : typeof e?.statusCode === "number"
        ? (e.statusCode as number)
        : undefined;

  if (status !== undefined) {
    // We got an HTTP response, so the request was delivered and understood.
    // A 4xx is our fault -- a malformed activity, an expired conversation, a bad
    // token -- and resending the identical request just reproduces it. 429 and
    // 5xx are the service asking us to come back later.
    return status === 429 || status >= 500;
  }

  // No HTTP status at all: the socket failed. This is the ECONNRESET case.
  const code = typeof e?.code === "string" ? e.code : undefined;
  const causeCode = typeof e?.cause?.code === "string" ? e.cause.code : undefined;
  if (code && RETRYABLE_NETWORK_CODES.includes(code)) return true;
  if (causeCode && RETRYABLE_NETWORK_CODES.includes(causeCode)) return true;

  // Last resort: some layers stringify the cause instead of preserving `code`.
  const message = typeof e?.message === "string" ? e.message : "";
  return RETRYABLE_NETWORK_CODES.some((c) => message.includes(c));
}

function backoffDelayMs(attempt: number): number {
  const exponential = Math.min(MAX_DELAY_MS, BASE_DELAY_MS * 2 ** (attempt - 1));
  // Full-ish jitter. Several conversations recovering from the same blip should
  // not retry in lockstep.
  return exponential * (0.5 + Math.random() * 0.5);
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

/**
 * Wraps a send function so transient network failures are retried with
 * exponential backoff. Generic over the activity type so this module stays
 * independent of the Teams SDK.
 */
export function withSendRetry<A>(
  send: (activity: A) => Promise<unknown>,
  log: ILogger
): (activity: A) => Promise<unknown> {
  return async (activity: A) => {
    for (let attempt = 1; ; attempt++) {
      try {
        return await send(activity);
      } catch (err) {
        if (attempt >= MAX_ATTEMPTS || !isRetryableSendError(err)) {
          if (attempt > 1) {
            log.error(`send: giving up after ${attempt} attempt(s)`, describeError(err));
          }
          throw err;
        }

        const delay = backoffDelayMs(attempt);
        log.warn(
          `send: attempt ${attempt}/${MAX_ATTEMPTS} failed, retrying in ${Math.round(delay)}ms`,
          describeError(err)
        );
        await sleep(delay);
      }
    }
  };
}
