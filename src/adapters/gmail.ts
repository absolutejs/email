import type {
  EmailDeltaResult,
  EmailFetch,
  EmailSubscriptionResult,
  NormalizedEmailMessage,
  TokenCredential,
} from "../types";
import {
  cleanEmail,
  directionFor,
  fetchJson,
  firstAddress,
  parseDate,
  stripHtml,
} from "../utils";

const GMAIL_API = "https://gmail.googleapis.com/gmail/v1/users/me";
const MAX_WATCH_LEASE_MS = 6 * 24 * 60 * 60 * 1000;

export type GmailHeader = {
  name?: string;
  value?: string;
};

export type GmailMessagePart = {
  body?: {
    data?: string;
  };
  mimeType?: string;
  parts?: GmailMessagePart[];
};

export type GmailMessage = {
  historyId?: string;
  id?: string;
  internalDate?: string;
  payload?: GmailMessagePart & {
    headers?: GmailHeader[];
  };
  snippet?: string;
  threadId?: string;
};

export type GmailHistoryResponse = {
  history?: {
    messages?: { id?: string; threadId?: string }[];
    messagesAdded?: { message?: { id?: string; threadId?: string } }[];
  }[];
  historyId?: string;
};

export type GmailWatchResponse = {
  expiration?: string;
  historyId?: string;
};

export type GmailClient = {
  getMessage: (id: string) => Promise<GmailMessage | null>;
  listHistory: (input: {
    cursor?: string | null;
  }) => Promise<EmailDeltaResult<{ id: string; threadId?: string }>>;
  watch: (topicName: string) => Promise<EmailSubscriptionResult & {
    cursor?: string | null;
  }>;
};

const decodeBase64Url = (value: string) =>
  Buffer.from(
    value.replace(/-/gu, "+").replace(/_/gu, "/"),
    "base64",
  ).toString("utf8");

const header = (message: GmailMessage, name: string) =>
  message.payload?.headers?.find(
    (item) => item.name?.toLowerCase() === name.toLowerCase(),
  )?.value ?? null;

const collectPartText = (part: GmailMessagePart | undefined): string | null => {
  if (!part) return null;
  if (part.mimeType === "text/plain" && part.body?.data) {
    return decodeBase64Url(part.body.data);
  }
  if (part.mimeType === "text/html" && part.body?.data) {
    return stripHtml(decodeBase64Url(part.body.data));
  }
  const nested = part.parts
    ?.map((child) => collectPartText(child))
    .find((value) => value && value.length > 0);

  return nested ?? null;
};

const parseAddress = (value: string | null | undefined) => {
  const raw = value ?? "";
  const bracket = raw.match(/^(.*)<([^>]+)>$/u);
  if (!bracket) return { address: cleanEmail(raw), name: null };

  return {
    address: cleanEmail(bracket[2]),
    name: bracket[1]?.replace(/^"|"$/gu, "").trim() || null,
  };
};

const parseAddressList = (value: string | null | undefined) =>
  (value ?? "")
    .split(",")
    .map(parseAddress)
    .filter((item) => item.address);

export const createGmailClient = (
  credential: TokenCredential,
  fetcher: EmailFetch = fetch,
): GmailClient => ({
  getMessage: async (id) => {
    const response = await fetchJson<GmailMessage>(
      `${GMAIL_API}/messages/${encodeURIComponent(id)}?format=full`,
      credential.accessToken,
      undefined,
      fetcher,
    );

    return response.ok ? response.body : null;
  },
  listHistory: async ({ cursor }) => {
    if (!cursor) return { cursor: null, expired: false, messages: [] };
    const params = new URLSearchParams({
      historyTypes: "messageAdded",
      startHistoryId: cursor,
    });
    const response = await fetchJson<GmailHistoryResponse>(
      `${GMAIL_API}/history?${params.toString()}`,
      credential.accessToken,
      undefined,
      fetcher,
    );
    if (response.status === 404) {
      return { cursor, expired: true, messages: [] };
    }
    if (!response.ok || !response.body) {
      return { cursor, expired: false, messages: [] };
    }
    const byId = new Map<string, { id: string; threadId?: string }>();
    for (const entry of response.body.history ?? []) {
      for (const added of entry.messagesAdded ?? []) {
        const id = added.message?.id;
        if (id) {
          byId.set(
            id,
            added.message?.threadId
              ? { id, threadId: added.message.threadId }
              : { id },
          );
        }
      }
      for (const message of entry.messages ?? []) {
        if (message.id) {
          byId.set(
            message.id,
            message.threadId
              ? { id: message.id, threadId: message.threadId }
              : { id: message.id },
          );
        }
      }
    }

    return {
      cursor: response.body.historyId ?? cursor,
      expired: false,
      messages: [...byId.values()],
    };
  },
  watch: async (topicName) => {
    const response = await fetchJson<GmailWatchResponse>(
      `${GMAIL_API}/watch`,
      credential.accessToken,
      {
        body: JSON.stringify({
          labelIds: ["INBOX", "SENT"],
          topicName,
        }),
        method: "POST",
      },
      fetcher,
    );
    if (!response.ok || !response.body) {
      return { cursor: null, expiration: null, id: null };
    }

    return {
      cursor: response.body.historyId ?? null,
      expiration: response.body.expiration
        ? new Date(Number(response.body.expiration))
        : new Date(Date.now() + MAX_WATCH_LEASE_MS),
      id: topicName,
    };
  },
});

export const gmailMessageToNormalized = (
  message: GmailMessage,
  input: { accountEmail: string },
): NormalizedEmailMessage | null => {
  if (!message.id) return null;
  const from = parseAddress(header(message, "from"));
  const to = parseAddressList(header(message, "to"));
  const bodyText = collectPartText(message.payload);

  return {
    accountEmail: cleanEmail(input.accountEmail),
    bodyText,
    direction: directionFor(input.accountEmail, from.address),
    from: from.address ? from : null,
    id: message.id,
    occurredAt: parseDate(message.internalDate),
    provider: "gmail",
    raw: message,
    snippet: message.snippet ?? null,
    subject: header(message, "subject"),
    threadId: message.threadId ?? null,
    to,
  };
};

export const gmailMessagesToNormalized = async (
  client: GmailClient,
  messages: { id: string }[],
  input: { accountEmail: string },
) => {
  const normalized: NormalizedEmailMessage[] = [];
  for (const item of messages) {
    const full = await client.getMessage(item.id);
    const row = full ? gmailMessageToNormalized(full, input) : null;
    if (row) normalized.push(row);
  }

  return normalized;
};

export const gmailCounterpartEmail = (message: NormalizedEmailMessage) =>
  message.direction === "outbound"
    ? firstAddress(message.to)
    : cleanEmail(message.from?.address);
