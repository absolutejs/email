import { gmailMessageToNormalized } from "./adapters/gmail";
import { describe, expect, test } from "bun:test";
import type { GmailClient } from "./adapters/gmail";
import { createMicrosoftGraphEmailClient } from "./adapters/microsoft";
import { manifest } from "./manifest";
import type { NormalizedEmailMessage } from "./types";
import {
  createGmailVerificationMessageLookup,
  createMicrosoftVerificationMessageLookup,
  EmailVerificationError,
  resolveEmailVerificationCode,
  retrieveEmailVerificationCode,
  type EmailVerificationProfile,
} from "./verification";
import { createImapVerificationMessageLookup } from "./verification-imap";

const NOW = new Date("2026-08-26T10:00:00.000Z");
const PROFILE: EmailVerificationProfile = {
  bodyMarkers: ["verification code"],
  id: "accounts-example-six-digit-v1",
  origins: ["https://accounts.example.com"],
  providers: ["gmail", "microsoft", "imap"],
  senderAddresses: ["security@example.com"],
  senderAuthentication: {
    allowedHeaderFromDomains: ["example.com"],
    trustedAuthservIds: ["mx.mailbox.example"],
  },
  subjectIncludesAny: ["sign in"],
};

const message = (
  overrides: Partial<NormalizedEmailMessage> = {},
): NormalizedEmailMessage => ({
  accountEmail: "member@example.net",
  authenticationResults: [
    "mx.mailbox.example; dmarc=pass header.from=example.com",
  ],
  bodyText: "Your verification code: 482193. It expires soon.",
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
      senderAuthenticated: true,
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

  test("requires one trusted, aligned DMARC pass", () => {
    const rejected = [
      message({ authenticationResults: [] }),
      message({
        authenticationResults: [
          "attacker.example; dmarc=pass header.from=example.com",
        ],
      }),
      message({
        authenticationResults: [
          "mx.mailbox.example; dmarc=fail header.from=example.com",
        ],
      }),
      message({
        authenticationResults: [
          "mx.mailbox.example; dmarc=pass header.from=attacker.example",
        ],
      }),
      message({
        authenticationResults: [
          "mx.mailbox.example; dmarc=pass header.from=example.com",
          "mx.mailbox.example; dmarc=pass header.from=example.com",
        ],
      }),
    ];

    for (const candidate of rejected) {
      expect(() => resolveEmailVerificationCode([candidate], query)).toThrow(
        new EmailVerificationError("no_match"),
      );
    }
  });

  test("rejects repeated codes, invalid dates, and missing correlation text", () => {
    expect(() =>
      resolveEmailVerificationCode(
        [
          message({
            bodyText: "Verification code: 482193. Verification code: 482193.",
          }),
        ],
        query,
      ),
    ).toThrow(new EmailVerificationError("ambiguous_match"));
    expect(() =>
      resolveEmailVerificationCode(
        [message({ occurredAt: new Date(Number.NaN) })],
        query,
      ),
    ).toThrow(new EmailVerificationError("no_match"));
    expect(() =>
      resolveEmailVerificationCode([message()], {
        ...query,
        requiredBodyText: ["challenge-that-is-not-present"],
      }),
    ).toThrow(new EmailVerificationError("no_match"));
  });

  test("bounds lookup windows, candidates, bodies, and provider queries", async () => {
    expect(() =>
      resolveEmailVerificationCode([message()], {
        ...query,
        notBefore: new Date(NOW.getTime() - 10 * 60_000 - 1),
      }),
    ).toThrow(new EmailVerificationError("invalid_profile"));
    expect(() =>
      resolveEmailVerificationCode([message(), message({ id: "message-2" })], {
        ...query,
        maxCandidates: 1,
      }),
    ).toThrow(new EmailVerificationError("candidate_limit"));

    let lookups = 0;
    await expect(
      retrieveEmailVerificationCode(
        {
          find: async () => {
            lookups += 1;
            return [];
          },
        },
        { ...query, accountEmail: "member@example.net after:0" },
      ),
    ).rejects.toEqual(new EmailVerificationError("invalid_profile"));
    expect(lookups).toBe(0);
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
            {
              name: "Authentication-Results",
              value: "mx.mailbox.example; dmarc=pass header.from=example.com",
            },
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
    expect(found[0]?.occurredAt).toEqual(NOW);
    expect(queries[0]).toContain("from:security@example.com");
    expect(queries[0]).toContain(`after:${query.notBefore.getTime() / 1000}`);
  });

  test("Microsoft uses Graph search with selected body fields and local normalization", async () => {
    const requestUrls: string[] = [];
    let prefer = "";
    const client = createMicrosoftGraphEmailClient(
      { accessToken: "secret-token" },
      (async (input, init) => {
        const requestUrl = String(input);
        requestUrls.push(requestUrl);
        prefer = new Headers(init?.headers).get("Prefer") ?? "";
        if (requestUrl.includes("/messages/graph-1?")) {
          return Response.json({
            body: { content: "Verification code: 482193" },
            from: { emailAddress: { address: "security@example.com" } },
            id: "graph-1",
            internetMessageHeaders: [
              {
                name: "Authentication-Results",
                value: "mx.mailbox.example; dmarc=pass header.from=example.com",
              },
            ],
            receivedDateTime: NOW.toISOString(),
            subject: "Sign in",
            toRecipients: [{ emailAddress: { address: "member@example.net" } }],
          });
        }
        return Response.json({
          value: [
            {
              id: "graph-1",
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
    expect(requestUrls[0]).toContain("%24search");
    expect(requestUrls[0]).toContain("from%3Asecurity%40example.com");
    expect(requestUrls[1]).toContain("internetMessageHeaders");
    expect(requestUrls.join("\n")).not.toContain("secret-token");
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
    expect(options).toEqual({ limit: 8, since: query.notBefore });
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

test("standalone layout is explicit, bounded, authenticated and rejects multiple codes", () => {
  const profile = {
    ...PROFILE,
    codeLayout: "standalone-after-marker" as const,
    bodyMarkers: ["different device"],
  };
  const bodyText =
    "It looks like you tried to sign in from a different device.\n\n482193\n\nIgnore this if it was not you.";
  expect(
    new TextDecoder().decode(
      resolveEmailVerificationCode([message({ bodyText })], {
        ...query,
        profile,
      }).bytes,
    ),
  ).toBe("482193");
  for (const body of [
    "482193\ndifferent device",
    "different device\n" + "x".repeat(513) + "\n482193",
    "different device\n482193\n591204",
    "different device\nvalue=482193",
    "different device\n4821937",
  ]) {
    expect(() =>
      resolveEmailVerificationCode([message({ bodyText: body })], {
        ...query,
        profile,
      }),
    ).toThrow();
  }
  expect(() =>
    resolveEmailVerificationCode(
      [message({ bodyText, authenticationResults: [] })],
      { ...query, profile },
    ),
  ).toThrow();
  expect(() =>
    resolveEmailVerificationCode([message({ bodyText })], {
      ...query,
      profile: { ...profile, codeLayout: "after-marker" },
    }),
  ).toThrow();
});

test("HTML-only Gmail preserves standalone code boundaries", () => {
  const body =
    "<p>Different device</p><div><strong>482193</strong></div><p>Expires shortly</p>";
  const normalized = gmailMessageToNormalized(
    {
      id: "html",
      internalDate: String(NOW.getTime()),
      payload: {
        mimeType: "text/html",
        headers: [
          { name: "from", value: "security@example.com" },
          { name: "to", value: "member@example.net" },
          { name: "subject", value: "Sign in to Example" },
          {
            name: "authentication-results",
            value: "mx.mailbox.example; dmarc=pass header.from=example.com",
          },
        ],
        body: { data: Buffer.from(body).toString("base64url") },
      },
    },
    { accountEmail: "member@example.net" },
  )!;
  expect(
    new TextDecoder().decode(
      resolveEmailVerificationCode([normalized], {
        ...query,
        profile: {
          ...PROFILE,
          bodyMarkers: ["Different device"],
          codeLayout: "standalone-after-marker",
        },
      }).bytes,
    ),
  ).toBe("482193");
});

describe("newest selection", () => {
  const older = message({
    bodyText: "Your verification code: 111111.",
    id: "older",
    occurredAt: new Date(NOW.getTime() - 20 * 60_000),
  });
  const newer = message({
    bodyText: "Your verification code: 222222.",
    id: "newer",
    occurredAt: new Date(NOW.getTime() - 60_000),
  });
  const wide = {
    ...query,
    notBefore: new Date(NOW.getTime() - 60 * 60_000),
  };

  test("unique selection still rejects several matches and wide windows", () => {
    expect(() => resolveEmailVerificationCode([older, newer], query)).toThrow();
    expect(() => resolveEmailVerificationCode([newer], wide)).toThrow(
      new EmailVerificationError("invalid_profile"),
    );
  });

  test("uses the most recent authenticated match across a wide window", () => {
    const result = resolveEmailVerificationCode([older, newer], {
      ...wide,
      selection: "newest",
    });
    expect(new TextDecoder().decode(result.bytes)).toBe("222222");
    expect(result.evidence.messageId).toBe("newer");
  });

  test("ignores a newer message that fails sender authentication", () => {
    const forged = message({
      authenticationResults: [
        "mx.mailbox.example; dmarc=fail header.from=example.com",
      ],
      bodyText: "Your verification code: 999999.",
      id: "forged",
      occurredAt: new Date(NOW.getTime() - 1000),
    });
    const result = resolveEmailVerificationCode([older, newer, forged], {
      ...wide,
      selection: "newest",
    });
    expect(new TextDecoder().decode(result.bytes)).toBe("222222");
  });

  test("two different codes at the same instant stay ambiguous", () => {
    const twin = message({
      bodyText: "Your verification code: 333333.",
      id: "twin",
      occurredAt: newer.occurredAt,
    });
    expect(() =>
      resolveEmailVerificationCode([newer, twin], {
        ...wide,
        selection: "newest",
      }),
    ).toThrow(new EmailVerificationError("ambiguous_match"));
  });

  test("rejects windows longer than 24 hours", () => {
    expect(() =>
      resolveEmailVerificationCode([newer], {
        ...query,
        notBefore: new Date(NOW.getTime() - 25 * 60 * 60_000),
        selection: "newest",
      }),
    ).toThrow(new EmailVerificationError("invalid_profile"));
  });
});

describe("Gmail newest lookup", () => {
  const gmailWith = (ids: string[], sizes: number[]) =>
    ({
      getMessage: async (id: string) => ({
        id,
        internalDate: String(NOW.getTime()),
        payload: {
          headers: [
            { name: "From", value: "security@example.com" },
            { name: "To", value: "member@example.net" },
            { name: "Subject", value: "Sign in" },
          ],
          mimeType: "text/plain",
          body: { data: Buffer.from("code").toString("base64url") },
        },
      }),
      listHistory: async () => ({ messages: [] }),
      searchMessages: async ({ maxResults }: { maxResults?: number }) => {
        sizes.push(maxResults ?? 0);
        return ids.slice(0, maxResults).map((id) => ({ id }));
      },
      watch: async () => ({}),
    }) as GmailClient;

  test("keeps the newest candidates instead of failing on a busy inbox", async () => {
    const sizes: number[] = [];
    const lookup = createGmailVerificationMessageLookup({
      accountEmail: "member@example.net",
      client: gmailWith(["a", "b", "c", "d"], sizes),
      maxCandidates: 2,
    });
    const found = await lookup.find({ ...query, selection: "newest" });
    expect(found.map((m) => m.id)).toEqual(["a", "b"]);
    expect(sizes).toEqual([2]);
  });

  test("unique selection still fails when candidates exceed the limit", async () => {
    const lookup = createGmailVerificationMessageLookup({
      accountEmail: "member@example.net",
      client: gmailWith(["a", "b", "c"], []),
      maxCandidates: 2,
    });
    await expect(lookup.find(query)).rejects.toEqual(
      new EmailVerificationError("candidate_limit"),
    );
  });
});
