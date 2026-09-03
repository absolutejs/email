// Resend Inbound: the `email.received` webhook carries METADATA ONLY (ids,
// addresses, subject) — the body and headers must be fetched from the
// received-emails API with the event's email_id.
import type { EmailFetch } from "../types";
import { parseMailbox, parseMailboxList } from "../inbound";
import type { InboundEmailMessage } from "../inbound";
import { stripHtml } from "../utils";

const RESEND_API = "https://api.resend.com";
const RECEIVED_EVENT = "email.received";

export type ResendInboundEventData = {
  attachments?: unknown[];
  bcc?: string[] | string;
  cc?: string[] | string;
  created_at?: string;
  email_id?: string;
  from?: string;
  headers?: Record<string, string> | { name?: string; value?: string }[];
  html?: string | null;
  message_id?: string;
  subject?: string;
  text?: string | null;
  to?: string[] | string;
};

export type ResendInboundWebhookBody = {
  created_at?: string;
  data?: ResendInboundEventData;
  type?: string;
};

type HeaderShape = NonNullable<ResendInboundEventData["headers"]>;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const asString = (value: unknown) => (typeof value === "string" ? value : "");

const asAddressList = (value: unknown) => {
  if (typeof value === "string") return value;
  if (Array.isArray(value)) {
    return value.filter((entry): entry is string => typeof entry === "string");
  }

  return null;
};

const readHeader = (headers: HeaderShape | undefined, name: string) => {
  if (!headers) return undefined;
  if (Array.isArray(headers)) {
    const found = headers.find(
      (header) => (header.name ?? "").toLowerCase() === name,
    );

    return found?.value?.trim() || undefined;
  }
  const entry = Object.entries(headers).find(
    ([key]) => key.toLowerCase() === name,
  );

  return entry?.[1]?.trim() || undefined;
};

const asHeaders = (value: unknown): HeaderShape | undefined => {
  if (Array.isArray(value)) {
    return value.filter(isRecord).map((entry) => ({
      name: asString(entry.name),
      value: asString(entry.value),
    }));
  }
  if (!isRecord(value)) return undefined;
  const map: Record<string, string> = {};
  for (const [key, entry] of Object.entries(value)) {
    if (typeof entry === "string") map[key] = entry;
  }

  return map;
};

const eventDataToInbound = (
  source: Record<string, unknown>,
): InboundEmailMessage | null => {
  const from = parseMailbox(asString(source.from));
  const to = parseMailboxList(asAddressList(source.to));
  const emailId = asString(source.email_id) || asString(source.id);
  if (!from || to.length === 0 || !emailId) return null;
  const headers = asHeaders(source.headers);
  const messageIdHeader =
    asString(source.message_id).trim() || readHeader(headers, "message-id");
  const inReplyTo = readHeader(headers, "in-reply-to");
  const references = readHeader(headers, "references");
  const text = typeof source.text === "string" ? source.text : undefined;
  const html = typeof source.html === "string" ? source.html : undefined;
  const createdAt = asString(source.created_at);
  const receivedAt = createdAt ? new Date(createdAt) : undefined;

  return {
    cc: parseMailboxList(asAddressList(source.cc)),
    emailId,
    from,
    ...(html === undefined ? {} : { html }),
    ...(inReplyTo === undefined ? {} : { inReplyTo }),
    ...(messageIdHeader ? { messageIdHeader } : {}),
    provider: "resend",
    ...(receivedAt && !Number.isNaN(receivedAt.getTime())
      ? { receivedAt }
      : {}),
    ...(references === undefined ? {} : { references }),
    subject: asString(source.subject).trim() || "(no subject)",
    ...(text === undefined ? {} : { text }),
    to,
  };
};

/** Normalize a Resend `email.received` webhook. Accepts the event envelope
 *  (`{ type, data: {…} }`) and the bare data object; any other event type is
 *  rejected with null. The result has no body — see
 *  `fetchResendReceivedEmail`. */
export const parseResendInboundWebhook = (
  body: unknown,
): InboundEmailMessage | null => {
  if (!isRecord(body)) return null;
  const type = asString(body.type);
  if (type && type !== RECEIVED_EVENT) return null;
  const source = isRecord(body.data) ? body.data : body;

  return eventDataToInbound(source);
};

/** GET /emails/receiving/:id — the full received message (text, html,
 *  headers). Returns null when the API rejects the request or the payload
 *  can't be normalized; a network failure never throws. */
export const fetchResendReceivedEmail = async (
  apiKey: string,
  emailId: string,
  fetcher: EmailFetch = fetch,
): Promise<InboundEmailMessage | null> => {
  const response = await fetcher(
    `${RESEND_API}/emails/receiving/${encodeURIComponent(emailId)}`,
    { headers: { Authorization: `Bearer ${apiKey}` } },
  ).catch(() => null);
  if (!response?.ok) return null;
  const payload: unknown = await response.json().catch(() => null);
  if (!isRecord(payload)) return null;
  const inbound = eventDataToInbound({ ...payload, email_id: emailId });
  if (!inbound) return null;
  // Text is what reply-stripping wants; derive it from HTML-only mail.
  if (inbound.text === undefined && inbound.html !== undefined) {
    return { ...inbound, text: stripHtml(inbound.html) };
  }

  return inbound;
};

/** Webhook event + fetched body in one call: the event is authoritative for
 *  routing (envelope recipients), the fetch supplies text/html/headers. */
export const resolveResendInboundEmail = async (
  apiKey: string,
  body: unknown,
  fetcher: EmailFetch = fetch,
): Promise<InboundEmailMessage | null> => {
  const event = parseResendInboundWebhook(body);
  if (!event) return null;
  if (event.text !== undefined || event.html !== undefined) return event;
  const received = await fetchResendReceivedEmail(
    apiKey,
    event.emailId,
    fetcher,
  );
  if (!received) return event;

  return {
    ...received,
    cc: event.cc.length > 0 ? event.cc : received.cc,
    from: event.from,
    subject: event.subject,
    to: event.to.length > 0 ? event.to : received.to,
  };
};
