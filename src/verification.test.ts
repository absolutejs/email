import { describe, expect, test } from "bun:test";
import type { GmailClient } from "./adapters/gmail";
import { createMicrosoftGraphEmailClient } from "./adapters/microsoft";
import { manifest } from "./manifest";
import type { NormalizedEmailMessage } from "./types";
import {
  createGmailVerificationMessageLookup,
  createImapVerificationMessageLookup,
  createMicrosoftVerificationMessageLookup,
  EmailVerificationError,
  resolveEmailVerificationCode,
  retrieveEmailVerificationCode,
  type EmailVerificationProfile,
} from "./verification";

const NOW = new Date("2026-08-26T10:00:00.000Z");
const PROFILE: EmailVerificationProfile = {
  bodyMarkers: ["verification code"],
  id: "accounts-example-six-digit-v1",
  origins: ["https://accounts.example.com"],
  providers: ["gmail", "microsoft", "imap"],
  senderAddresses: ["security@example.com"],
  subjectIncludesAny: ["sign in"],
};

const message = (
  overrides: Partial<NormalizedEmailMessage> = {},
): NormalizedEmailMessage => ({
  accountEmail: "member@example.net",
  bodyText: "Your verification code is: 482193. It expires soon.",
  direction: "inbound",
  from: { address: "security@example.com" },
  id: "message-1",
  occurredAt: NOW,
  provider: "gmail",
  subject: "Sign in to Example",
  to: [{ address: "member@example.net" }],
  ...overrides,
});

const query = {
  accountEmail: "member@example.net",
  expectedOrigin: "https://accounts.example.com",
  notAfter: new Date(NOW.getTime() + 30_000),
  notBefore: new Date(NOW.getTime() - 30_000),
  profile: PROFILE,
};

describe("deterministic email verification parsing", () => {
  test("returns mutable bytes and non-secret evidence for one exact match", () => {
    const result = resolveEmailVerificationCode([message()], query);

    expect(new TextDecoder().decode(result.bytes)).toBe("482193");
    expect(result.evidence).toEqual({
      matchedAt: NOW.getTime(),
      messageId: "message-1",
      parserId: PROFILE.id,
      provider: "gmail",
    });
    expect(JSON.stringify(result.evidence)).not.toContain("482193");
  });

  test("rejects sender, subject, account, provider, direction, and time mismatches", () => {
    const mismatches = [
      message({ from: { address: "attacker@example.com" } }),
      message({ subject: "Your receipt" }),
      message({ accountEmail: "someone-else@example.net" }),
      message({ provider: "other" }),
      message({ direction: "outbound" }),
      message({ occurredAt: new Date(NOW.getTime() - 60_000) }),
    ];

    for (const candidate of mismatches) {
      expect(() => resolveEmailVerificationCode([candidate], query)).toThrow(
        new EmailVerificationError("no_match"),
      );
    }
  });

  test("fails closed when messages or marker-bound codes are ambiguous", () => {
    expect(() =>
      resolveEmailVerificationCode(
        [message(), message({ id: "message-2" })],
        query,
      ),
    ).toThrow(new EmailVerificationError("ambiguous_match"));
    expect(() =>
      resolveEmailVerificationCode(
        [
          message({
            bodyText:
              "Verification code: 482193. Previous verification code: 917204.",
          }),
        ],
        query,
      ),
    ).toThrow(new EmailVerificationError("ambiguous_match"));
  });

  test("does not accept an unbound origin or an unmarked number", () => {
    expect(() =>
      resolveEmailVerificationCode([message()], {
        ...query,
        expectedOrigin: "https://attacker.example",
      }),
    ).toThrow(new EmailVerificationError("invalid_profile"));
    expect(() =>
      resolveEmailVerificationCode(
        [message({ bodyText: "Order 482193 is ready." })],
        query,
      ),
    ).toThrow(new EmailVerificationError("no_match"));
  });

  test("maps lookup failures to a safe error without leaking provider text", async () => {
    await expect(
      retrieveEmailVerificationCode(
        {
          find: () => {
            throw new Error("provider response contained 482193");
          },
        },
        query,
      ),
    ).rejects.toEqual(new EmailVerificationError("lookup_failed"));
  });
});

describe("provider lookups", () => {
  test("Gmail narrows by exact sender and epoch window before full fetch", async () => {
    const queries: string[] = [];
    const gmail = {
      getMessage: async () => ({
        id: "gmail-1",
        internalDate: String(NOW.getTime()),
        payload: {
          headers: [
            { name: "From", value: "security@example.com" },
            { name: "To", value: "member@example.net" },
            { name: "Subject", value: "Sign in" },
          ],
          mimeType: "text/plain",
          body: {
            data: Buffer.from("Verification code: 482193").toString(
              "base64url",
            ),
          },
        },
      }),
      listHistory: async () => ({ messages: [] }),
      searchMessages: async ({ query: value }: { query: string }) => {
        queries.push(value);
        return [{ id: "gmail-1" }];
      },
      watch: async () => ({}),
    } as GmailClient;
    const lookup = createGmailVerificationMessageLookup({
      accountEmail: "member@example.net",
      client: gmail,
    });

    const found = await lookup.find(query);
    expect(found).toHaveLength(1);
    expect(queries[0]).toContain("from:security@example.com");
    expect(queries[0]).toContain(`after:${query.notBefore.getTime() / 1000}`);
  });

  test("Microsoft uses Graph search with selected body fields and local normalization", async () => {
    let requestUrl = "";
    let prefer = "";
    const client = createMicrosoftGraphEmailClient(
      { accessToken: "secret-token" },
      (async (input, init) => {
        requestUrl = String(input);
        prefer = new Headers(init?.headers).get("Prefer") ?? "";
        return Response.json({
          value: [
            {
              body: { content: "Verification code: 482193" },
              from: {
                emailAddress: { address: "security@example.com" },
              },
              id: "graph-1",
              receivedDateTime: NOW.toISOString(),
              subject: "Sign in",
              toRecipients: [
                { emailAddress: { address: "member@example.net" } },
              ],
            },
          ],
        });
      }) as typeof fetch,
    );
    const lookup = createMicrosoftVerificationMessageLookup({
      accountEmail: "member@example.net",
      client,
    });

    const found = await lookup.find(query);
    expect(found).toHaveLength(1);
    expect(requestUrl).toContain("%24search");
    expect(requestUrl).toContain("from%3Asecurity%40example.com");
    expect(requestUrl).toContain("body");
    expect(requestUrl).not.toContain("secret-token");
    expect(prefer).toContain("text");
  });

  test("IMAP requests only the bounded recent window", async () => {
    let options: unknown;
    const lookup = createImapVerificationMessageLookup({
      config: {
        accountEmail: "member@example.net",
        auth: { pass: "app-password", user: "member@example.net" },
        host: "imap.example.net",
      },
      fetch: async (_config, requested) => {
        options = requested;
        return { messages: [message({ provider: "imap" })] };
      },
      maxCandidates: 7,
    });

    expect(await lookup.find(query)).toHaveLength(1);
    expect(options).toEqual({ limit: 7, since: query.notBefore });
  });
});

test("verification retrieval is not registered as a model-facing email tool", () => {
  expect(Object.keys(manifest.tools ?? {})).not.toContain("verification_code");
  expect(Object.keys(manifest.tools ?? {})).not.toContain(
    "retrieve_verification_code",
  );
});

test("source contracts use type aliases, not interfaces", async () => {
  const source = await Bun.file(
    new URL("./verification.ts", import.meta.url),
  ).text();
  expect(source).not.toMatch(/\binterface\s+[A-Za-z_$]/u);
});
