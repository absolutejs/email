import { expect, test } from "bun:test";
import { parseGmailPubSubWebhook } from "./gmail";

const notification = (value: unknown) => ({
  message: { data: Buffer.from(JSON.stringify(value)).toString("base64") },
});
const cases = (values: unknown[]): [unknown][] =>
  values.map((value) => [value]);
const emailAddress = "Owner@Example.com";

test.each(cases([0, 9876543210, Number.MAX_SAFE_INTEGER]))(
  "normalizes safe numeric cursor %p to a string",
  (historyId) => {
    expect(
      parseGmailPubSubWebhook(notification({ emailAddress, historyId })),
    ).toEqual({ emailAddress, historyId: String(historyId) });
  },
);

test.each(cases(["0", "9876543210", "9007199254740993123"]))(
  "preserves string cursor %p exactly",
  (historyId) => {
    expect(
      parseGmailPubSubWebhook(notification({ emailAddress, historyId })),
    ).toEqual({ emailAddress, historyId });
  },
);

test.each(
  cases([
    undefined,
    null,
    -1,
    1.5,
    Number.MAX_SAFE_INTEGER + 1,
    "",
    "invalid",
    "1.5",
    "-1",
    true,
    {},
    [],
  ]),
)(
  "omits unusable cursor %p without discarding the notification",
  (historyId) => {
    const result = parseGmailPubSubWebhook(
      notification({ emailAddress, historyId }),
    );
    expect(result).toEqual({ emailAddress });
    expect(Object.hasOwn(result!, "historyId")).toBe(false);
  },
);

test.each(cases([null, [], {}, { emailAddress: 123 }, { emailAddress: " " }]))(
  "rejects invalid decoded payload %p",
  (value) => expect(parseGmailPubSubWebhook(notification(value))).toBeNull(),
);

test.each(
  cases([
    null,
    undefined,
    [],
    {},
    { message: null },
    { message: { data: 123 } },
    { message: { data: "not-json" } },
  ]),
)("rejects malformed envelope %p without throwing", (body) =>
  expect(parseGmailPubSubWebhook(body)).toBeNull(),
);

test("returns only validated fields", () => {
  expect(
    parseGmailPubSubWebhook(
      notification({ emailAddress, historyId: "42", unexpected: true }),
    ),
  ).toEqual({ emailAddress, historyId: "42" });
});
