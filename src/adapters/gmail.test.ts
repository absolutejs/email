import { expect, test } from "bun:test";
import { createGmailClient } from "./gmail";
import type { EmailFetch } from "../types";
const client = (responses: Response[]) =>
  createGmailClient({ accessToken: "fixture" }, (async () =>
    responses.shift()!) as EmailFetch);
test("history never returns an advanced cursor after partial failure or page exhaustion", async () => {
  const first = () =>
    Response.json({
      history: [{ messagesAdded: [{ message: { id: "one" } }] }],
      historyId: "9999999999999999999",
      nextPageToken: "two",
    });
  await expect(
    client([first(), new Response("", { status: 503 })]).listHistory({
      cursor: "1",
    }),
  ).rejects.toThrow("failed");
  await expect(
    client([first()]).listHistory({ cursor: "1", maxPages: 1 }),
  ).rejects.toThrow("incomplete");
  await expect(
    client([Response.json({ history: [] })]).listHistory({ cursor: "1" }),
  ).rejects.toThrow("valid cursor");
});
test("history preserves exact large cursors, deduplicates and reports expiry", async () => {
  const result = await client([
    Response.json({
      history: [{ messages: [{ id: "one" }] }],
      historyId: "20",
      nextPageToken: "two",
    }),
    Response.json({
      history: [
        {
          messagesAdded: [
            { message: { id: "one" } },
            { message: { id: "two", threadId: "thread" } },
          ],
        },
      ],
      historyId: "9999999999999999999",
    }),
  ]).listHistory({ cursor: "1" });
  expect(result).toEqual({
    cursor: "9999999999999999999",
    expired: false,
    messages: [{ id: "one" }, { id: "two", threadId: "thread" }],
  });
  expect(
    await client([new Response("", { status: 404 })]).listHistory({
      cursor: "1",
    }),
  ).toEqual({ cursor: "1", expired: true, messages: [] });
  await expect(
    client([]).listHistory({ cursor: "1", maxPages: 101 }),
  ).rejects.toThrow("maxPages");
});
test("failed reads and watch calls are not empty successes; deleted messages remain nullable", async () => {
  await expect(
    client([new Response("", { status: 401 })]).getMessage("id"),
  ).rejects.toThrow("failed");
  expect(
    await client([new Response("", { status: 404 })]).getMessage("id"),
  ).toBeNull();
  await expect(
    client([new Response("", { status: 500 })]).searchMessages({
      query: "job",
    }),
  ).rejects.toThrow("failed");
  await expect(
    client([new Response("", { status: 403 })]).watch("topic"),
  ).rejects.toThrow("failed");
});
