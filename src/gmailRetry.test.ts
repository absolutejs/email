import { expect, test } from "bun:test";
import { createRetryingGmailReadFetch } from "./gmailRetry";
const retry = (
  fetcher: (input: RequestInfo | URL, init?: RequestInit) => Promise<Response>,
  options?: Parameters<typeof createRetryingGmailReadFetch>[1],
) =>
  createRetryingGmailReadFetch(
    Object.assign(fetcher, { preconnect: fetch.preconnect }),
    options,
  );
const url = "https://gmail.googleapis.com/gmail/v1/users/me/messages/fixture";
const limited = () =>
  Response.json(
    { error: { errors: [{ reason: "rateLimitExceeded" }] } },
    { status: 403 },
  );
test("retries Gmail read throttles with bounded backoff and preserves final errors", async () => {
  let calls = 0;
  const delays: number[] = [];
  const read = retry(
    async () => (++calls === 3 ? Response.json({ id: "ok" }) : limited()),
    {
      wait: async (ms) => {
        delays.push(ms);
      },
    },
  );
  expect(await (await read(url)).json()).toEqual({ id: "ok" });
  expect(calls).toBe(3);
  expect(delays[0]).toBeGreaterThanOrEqual(1000);
  expect(delays[1]).toBeGreaterThanOrEqual(2000);
  calls = 0;
  const bounded = retry(
    async () => {
      calls++;
      return limited();
    },
    { maxAttempts: 2, wait: async () => {} },
  );
  expect((await bounded(url)).status).toBe(403);
  expect(calls).toBe(2);
});
test("never retries sends, other hosts, permission errors or long Retry-After", async () => {
  for (const [target, method, response] of [
    [url, "POST", limited()],
    ["https://example.com/mail", "GET", limited()],
    [
      url,
      "GET",
      Response.json(
        { error: { errors: [{ reason: "domainPolicy" }] } },
        { status: 403 },
      ),
    ],
    [url, "GET", new Response("", { status: 401 })],
    [
      url,
      "GET",
      new Response("", { status: 429, headers: { "retry-after": "120" } }),
    ],
  ] as const) {
    let calls = 0;
    const read = retry(
      async () => {
        calls++;
        return response;
      },
      {
        wait: async () => {
          throw new Error("unexpected retry");
        },
      },
    );
    expect(await read(target, { method })).toBe(response);
    expect(calls).toBe(1);
  }
});
test("cancellation prevents another attempt", async () => {
  const controller = new AbortController();
  let calls = 0;
  const read = retry(
    async () => {
      calls++;
      return limited();
    },
    {
      wait: async () => {
        controller.abort();
      },
    },
  );
  await expect(read(url, { signal: controller.signal })).rejects.toThrow();
  expect(calls).toBe(1);
});
