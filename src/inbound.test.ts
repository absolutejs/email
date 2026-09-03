import { describe, expect, test } from "bun:test";
import {
  parseMailbox,
  parseMailboxList,
  replyPreview,
  stripQuotedReply,
} from "./inbound";

describe("parseMailbox", () => {
  test("reads display-name and bare forms", () => {
    expect(parseMailbox("Jane Founder <Jane@Acme.com>")).toEqual({
      address: "jane@acme.com",
      name: "Jane Founder",
    });
    expect(parseMailbox('"Doe, Jane" <jane@acme.com>')).toEqual({
      address: "jane@acme.com",
      name: "Doe, Jane",
    });
    expect(parseMailbox("jane@acme.com")).toEqual({
      address: "jane@acme.com",
      name: null,
    });
    expect(parseMailbox("")).toBeNull();
    expect(parseMailbox("not an address")).toBeNull();
  });

  test("parses header lists and arrays", () => {
    expect(parseMailboxList('a@x.com, "Doe, Jane" <j@y.com>, junk')).toEqual([
      { address: "a@x.com", name: null },
      { address: "j@y.com", name: "Doe, Jane" },
    ]);
    expect(parseMailboxList(["r+abc@reply.example", null, ""])).toEqual([
      { address: "r+abc@reply.example", name: null },
    ]);
  });
});

describe("stripQuotedReply", () => {
  test("keeps only the words above a Gmail attribution", () => {
    const text = [
      "Sounds great — Tuesday works.",
      "",
      "Dan",
      "",
      "On Tue, Sep 1, 2026 at 3:10 PM Jane Founder <",
      "jane@acme.com> wrote:",
      "",
      "> Hi Dan, would you be open to a partnership?",
    ].join("\n");

    expect(stripQuotedReply(text)).toBe("Sounds great — Tuesday works.\n\nDan");
  });

  test("cuts at quoted lines, signatures and Outlook blocks", () => {
    expect(stripQuotedReply("Yes please\n> earlier")).toBe("Yes please");
    expect(stripQuotedReply("Yes please\n-- \nDan Smith\nCEO")).toBe(
      "Yes please",
    );
    expect(
      stripQuotedReply(
        "Yes please\n\n-----Original Message-----\nFrom: Jane\nSent: Monday",
      ),
    ).toBe("Yes please");
    expect(
      stripQuotedReply(
        "Yes please\r\n\r\nFrom: Jane <j@y.com>\r\nSent: Monday",
      ),
    ).toBe("Yes please");
    expect(stripQuotedReply("Yes please\n\nSent from my iPhone")).toBe(
      "Yes please",
    );
  });

  test("returns an empty string for pure quoted history and caps length", () => {
    expect(stripQuotedReply("> nothing new")).toBe("");
    expect(stripQuotedReply(null)).toBe("");
    expect(stripQuotedReply("abcdefghij", { maxLength: 4 })).toBe("abcd");
  });

  test("does not mistake a sentence starting with On for an attribution", () => {
    expect(stripQuotedReply("On balance I think yes.\nLet's talk.")).toBe(
      "On balance I think yes.\nLet's talk.",
    );
  });
});

describe("replyPreview", () => {
  test("collapses whitespace and ellipsizes", () => {
    expect(replyPreview("Sounds\n\ngreat  —  Tuesday works.")).toBe(
      "Sounds great — Tuesday works.",
    );
    expect(replyPreview("abcdefghijklmnop", 8)).toBe("abcdefg…");
  });
});
