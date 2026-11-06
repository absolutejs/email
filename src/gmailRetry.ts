import type { EmailFetch } from "./types";

export type GmailReadRetryOptions = {
  maxAttempts?: number;
  /** Host-owned clock/scheduler hook, useful for deterministic tests. */
  wait?: (milliseconds: number, signal?: AbortSignal | null) => Promise<void>;
};

const waitForRetry = (milliseconds: number, signal?: AbortSignal | null) =>
  new Promise<void>((resolve, reject) => {
    signal?.throwIfAborted();
    const abort = () => {
      clearTimeout(timer);
      signal?.removeEventListener("abort", abort);
      reject(signal?.reason);
    };
    const timer = setTimeout(() => {
      signal?.removeEventListener("abort", abort);
      resolve();
    }, milliseconds);
    signal?.addEventListener("abort", abort, { once: true });
  });

/** Retry only Gmail GET reads through the host's already-authorized transport.
 * Other requests (including sending) are passed through exactly once.
 * Exhaustion returns the final error response for the caller to handle.
 */
export const createRetryingGmailReadFetch = (
  fetcher: EmailFetch,
  options: GmailReadRetryOptions = {},
): EmailFetch => {
  const maxAttempts = options.maxAttempts ?? 6;
  if (!Number.isInteger(maxAttempts) || maxAttempts < 1 || maxAttempts > 8)
    throw new Error("Gmail maxAttempts must be between 1 and 8");
  const wrapped = async (
    input: Parameters<EmailFetch>[0],
    init?: Parameters<EmailFetch>[1],
  ) => {
    const url = new URL(input instanceof Request ? input.url : String(input));
    const method = (
      init?.method ?? (input instanceof Request ? input.method : "GET")
    ).toUpperCase();
    if (
      url.protocol !== "https:" ||
      url.hostname !== "gmail.googleapis.com" ||
      method !== "GET"
    )
      return fetcher(input, init);
    const signal =
      init?.signal ?? (input instanceof Request ? input.signal : undefined);
    for (let attempt = 0; ; attempt++) {
      signal?.throwIfAborted();
      const response = await fetcher(input, init);
      if (attempt + 1 >= maxAttempts) return response;
      let retry = [429, 500, 502, 503, 504].includes(response.status);
      if (response.status === 403) {
        const body = await response
          .clone()
          .json()
          .catch(() => null);
        retry =
          Array.isArray(body?.error?.errors) &&
          body.error.errors.some(
            (entry: { reason?: unknown }) =>
              entry?.reason === "rateLimitExceeded" ||
              entry?.reason === "userRateLimitExceeded",
          );
      }
      if (!retry) return response;
      const header = response.headers.get("retry-after");
      const requested =
        header === null
          ? 0
          : /^\d+$/.test(header)
            ? Number(header) * 1000
            : Date.parse(header) - Date.now();
      // Do not retry before a server-requested delay that exceeds our bounded wait.
      if (Number.isFinite(requested) && requested > 30000) return response;
      const delay = Math.min(
        30000,
        Math.max(
          Number.isFinite(requested) ? requested : 0,
          1000 * 2 ** attempt + Math.floor(Math.random() * 250),
        ),
      );
      await response.body?.cancel();
      await (options.wait ?? waitForRetry)(delay, signal);
    }
  };
  return Object.assign(wrapped, { preconnect: fetcher.preconnect });
};
