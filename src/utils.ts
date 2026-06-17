import type { EmailAddress, EmailFetch, FetchJsonResult } from "./types";

export const cleanEmail = (value: string | null | undefined) =>
  (value ?? "").trim().toLowerCase();

export const stripHtml = (value: string | null | undefined) =>
  (value ?? "")
    .replace(/<style[\s\S]*?<\/style>/giu, " ")
    .replace(/<script[\s\S]*?<\/script>/giu, " ")
    .replace(/<[^>]+>/gu, " ")
    .replace(/&nbsp;/giu, " ")
    .replace(/&amp;/giu, "&")
    .replace(/&lt;/giu, "<")
    .replace(/&gt;/giu, ">")
    .replace(/\s+/gu, " ")
    .trim();

export const parseDate = (value: string | number | undefined) => {
  if (value === undefined) return new Date();
  const parsed = new Date(value);

  return Number.isNaN(parsed.getTime()) ? new Date() : parsed;
};

export const directionFor = (accountEmail: string, from?: string | null) =>
  cleanEmail(from) === cleanEmail(accountEmail) ? "outbound" : "inbound";

export const firstAddress = (addresses: EmailAddress[]) =>
  cleanEmail(addresses[0]?.address);

export const fetchJson = async <T>(
  input: string,
  accessToken: string,
  init?: RequestInit,
  fetcher: EmailFetch = fetch,
): Promise<FetchJsonResult<T>> => {
  const response = await fetcher(input, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...(init?.headers ?? {}),
    },
  }).catch(() => null);
  if (!response) return { body: null, ok: false, status: 0 };
  const body = (await response.json().catch(() => null)) as T | null;

  return { body, ok: response.ok, status: response.status };
};
