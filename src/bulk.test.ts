import { describe, expect, test } from "bun:test";
import { bulkSignalsFromHeaders, isBulkMail } from "./bulk";

const gmailShape = (name: string, value: string) => [{ name, value }];

describe("bulk mail signals", () => {
  test("a mailing list is reported on its own", () => {
    const signals = bulkSignalsFromHeaders(
      gmailShape("List-Unsubscribe", "<https://x.com/u>"),
    );
    expect(signals).toEqual({
      autoSubmitted: false,
      bulkPrecedence: false,
      mailingList: true,
    });
  });

  test("List-Id and Feedback-ID also mark a mailing system", () => {
    expect(bulkSignalsFromHeaders(gmailShape("List-Id", "<n.x.com>")).mailingList).toBe(true);
    expect(bulkSignalsFromHeaders(gmailShape("Feedback-ID", "a:b")).mailingList).toBe(true);
  });

  test("Auto-Submitted: no is NOT an auto-reply", () => {
    expect(bulkSignalsFromHeaders(gmailShape("Auto-Submitted", "no")).autoSubmitted).toBe(false);
    expect(
      bulkSignalsFromHeaders(gmailShape("Auto-Submitted", "auto-generated")).autoSubmitted,
    ).toBe(true);
  });

  test("precedence only counts for the bulk-ish values", () => {
    expect(bulkSignalsFromHeaders(gmailShape("Precedence", "bulk")).bulkPrecedence).toBe(true);
    expect(bulkSignalsFromHeaders(gmailShape("Precedence", "list")).bulkPrecedence).toBe(true);
    expect(bulkSignalsFromHeaders(gmailShape("Precedence", "urgent")).bulkPrecedence).toBe(false);
  });

  test("header names match case-insensitively", () => {
    expect(bulkSignalsFromHeaders(gmailShape("list-unsubscribe", "x")).mailingList).toBe(true);
    expect(bulkSignalsFromHeaders(gmailShape("LIST-UNSUBSCRIBE", "x")).mailingList).toBe(true);
  });

  test("accepts a plain header map as well as the array shape", () => {
    expect(bulkSignalsFromHeaders({ "List-Unsubscribe": "x" }).mailingList).toBe(true);
  });

  test("a personal message trips nothing", () => {
    const signals = bulkSignalsFromHeaders(gmailShape("Subject", "lunch?"));
    expect(isBulkMail(signals)).toBe(false);
  });

  test("missing headers are not an error", () => {
    expect(isBulkMail(bulkSignalsFromHeaders(null))).toBe(false);
    expect(isBulkMail(bulkSignalsFromHeaders(undefined))).toBe(false);
    expect(isBulkMail(bulkSignalsFromHeaders([]))).toBe(false);
  });
});
