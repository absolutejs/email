import { describe, expect, test } from "bun:test";
import type { EmailFetch } from "../types";
import {
  fetchResendReceivedEmail,
  parseResendInboundWebhook,
  resolveResendInboundEmail,
} from "./resend";

const envelope = {
  created_at: "2026-09-02T10:00:00.000Z",
  data: {
    created_at: "2026-09-02T09:59:58.000Z",
    email_id: "rcv_123",
    from: "Dan Partner <dan@trinity.com>",
    message_id: "<abc@mail.trinity.com>",
    subject: "Re: Partnership?",
    to: ["r+tok123@email.onspark.com"],
  },
  type: "email.received",
};

const fetchStub =
  (status: number, payload: unknown): EmailFetch =>
  async () =>
    new Response(JSON.stringify(payload), {
      headers: { "content-type": "application/json" },
      status,
    });

describe("parseResendInboundWebhook", () => {
  test("normalizes the event envelope", () => {
    const parsed = parseResendInboundWebhook(envelope);

    expect(parsed).toEqual({
      cc: [],
      emailId: "rcv_123",
      from: { address: "dan@trinity.com", name: "Dan Partner" },
      messageIdHeader: "<abc@mail.trinity.com>",
      provider: "resend",
      receivedAt: new Date("2026-09-02T09:59:58.000Z"),
      subject: "Re: Partnership?",
      to: [{ address: "r+tok123@email.onspark.com", name: null }],
    });
  });

  test("accepts the bare data shape and inline text", () => {
    const parsed = parseResendInboundWebhook({
      email_id: "rcv_9",
      from: "dan@trinity.com",
      subject: "",
      text: "Yes",
      to: "r+tok@email.onspark.com",
    });

    expect(parsed?.subject).toBe("(no subject)");
    expect(parsed?.text).toBe("Yes");
    expect(parsed?.to[0]?.address).toBe("r+tok@email.onspark.com");
  });

  test("rejects other event types and malformed bodies", () => {
    expect(
      parseResendInboundWebhook({ ...envelope, type: "email.delivered" }),
    ).toBeNull();
    expect(parseResendInboundWebhook(null)).toBeNull();
    expect(parseResendInboundWebhook({ data: { from: "x@y.com" } })).toBeNull();
  });
});

describe("fetchResendReceivedEmail", () => {
  test("reads body + threading headers from the receiving API", async () => {
    const received = await fetchResendReceivedEmail(
      "re_key",
      "rcv_123",
      fetchStub(200, {
        from: "dan@trinity.com",
        headers: [
          { name: "In-Reply-To", value: "<out@onspark.com>" },
          { name: "References", value: "<out@onspark.com>" },
        ],
        html: "<p>Yes, <b>Tuesday</b></p>",
        id: "rcv_123",
        subject: "Re: Partnership?",
        to: ["r+tok123@email.onspark.com"],
      }),
    );

    expect(received?.inReplyTo).toBe("<out@onspark.com>");
    expect(received?.references).toBe("<out@onspark.com>");
    expect(received?.text).toBe("Yes, Tuesday");
    expect(received?.html).toContain("<b>");
  });

  test("returns null on a rejected request", async () => {
    expect(
      await fetchResendReceivedEmail("re_key", "rcv_x", fetchStub(404, {})),
    ).toBeNull();
  });
});

describe("resolveResendInboundEmail", () => {
  test("keeps the event's routing and takes the fetched body", async () => {
    const email = await resolveResendInboundEmail(
      "re_key",
      envelope,
      fetchStub(200, {
        from: "someone-else@trinity.com",
        id: "rcv_123",
        subject: "different",
        text: "Sounds great\n\n> quoted",
        to: ["r+tok123@email.onspark.com"],
      }),
    );

    expect(email?.from.address).toBe("dan@trinity.com");
    expect(email?.subject).toBe("Re: Partnership?");
    expect(email?.text).toBe("Sounds great\n\n> quoted");
  });
});
