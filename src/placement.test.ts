import { describe, expect, test } from "bun:test";
import type { GmailClient } from "./adapters/gmail";
import {
  findGmailMessagePlacement,
  gmailPlacementFromLabels,
} from "./placement";

describe("email placement", () => {
  test("classifies Gmail system labels", () => {
    expect(gmailPlacementFromLabels(["INBOX", "IMPORTANT"])).toBe("inbox");
    expect(gmailPlacementFromLabels(["SPAM"])).toBe("spam");
    expect(gmailPlacementFromLabels(["TRASH"])).toBe("trash");
    expect(gmailPlacementFromLabels(["CATEGORY_PROMOTIONS"])).toBe("other");
  });

  test("search includes Spam and returns the matched placement", async () => {
    const client = {
      getMessage: async () => ({ id: "message-1", labelIds: ["SPAM"] }),
      searchMessages: async () => [{ id: "message-1" }],
    } as Pick<GmailClient, "getMessage" | "searchMessages"> as GmailClient;

    expect(
      await findGmailMessagePlacement(client, {
        query: "subject:onspark-deliverability-canary-123",
      }),
    ).toEqual({ messageId: "message-1", placement: "spam" });
  });

  test("reports missing when no canary matches", async () => {
    const client = {
      getMessage: async () => null,
      searchMessages: async () => [],
    } as Pick<GmailClient, "getMessage" | "searchMessages"> as GmailClient;

    expect(
      await findGmailMessagePlacement(client, { query: "canary-token" }),
    ).toEqual({ messageId: null, placement: "missing" });
  });
});
