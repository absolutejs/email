// Inbound mail — a message a receiving provider (Resend, SES, Postmark…)
// delivered to one of the host's addresses. The provider-specific webhook
// parsers in ./webhooks normalize to this shape so a host can route by
// recipient (signed reply aliases, plus-addressing) and read the human part
// of a reply without knowing which provider carried it.
import type { EmailAddress } from "./types";

export type InboundEmailMessage = {
  cc: EmailAddress[];
  /** The provider's id for the received message — the idempotency key. */
  emailId: string;
  from: EmailAddress;
  html?: string;
  /** RFC 5322 `In-Reply-To` when the provider exposes headers. */
  inReplyTo?: string;
  /** RFC 5322 `Message-ID` of the received message. */
  messageIdHeader?: string;
  provider: "resend" | (string & {});
  receivedAt?: Date;
  /** RFC 5322 `References` chain, space-separated as received. */
  references?: string;
  subject: string;
  text?: string;
  to: EmailAddress[];
};

const MAILBOX_PATTERN = /^\s*(?:"?([^"<]*?)"?\s*)?<([^<>\s]+@[^<>\s]+)>\s*$/u;
const BARE_ADDRESS_PATTERN = /^[^\s<>,;"']+@[^\s<>,;"']+$/u;

/** Parse one RFC 5322 mailbox (`Jane <jane@x.com>` or `jane@x.com`). */
export const parseMailbox = (
  value: string | null | undefined,
): EmailAddress | null => {
  const raw = (value ?? "").trim();
  if (!raw) return null;
  const bracketed = raw.match(MAILBOX_PATTERN);
  if (bracketed) {
    const [, name, address] = bracketed;
    if (!address) return null;
    const cleanName = (name ?? "").trim();

    return {
      address: address.toLowerCase(),
      name: cleanName || null,
    };
  }
  if (BARE_ADDRESS_PATTERN.test(raw)) {
    return { address: raw.toLowerCase(), name: null };
  }
  const embedded = raw.match(/[^\s<>,;"']+@[^\s<>,;"']+/u);

  return embedded ? { address: embedded[0].toLowerCase(), name: null } : null;
};

/** Parse a header-style list (`a@x.com, Jane <j@y.com>`) or an array of
 *  mailboxes, dropping anything that isn't an address. */
export const parseMailboxList = (
  value: string | (string | null | undefined)[] | null | undefined,
): EmailAddress[] => {
  const entries = Array.isArray(value)
    ? value
    : (value ?? "").split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/u);

  return entries
    .map((entry) => parseMailbox(entry))
    .filter((entry): entry is EmailAddress => entry !== null);
};

// Where the human's words stop and the quoted history / signature begins.
// Each pattern is anchored to a line start; the earliest match wins.
const QUOTE_MARKERS: RegExp[] = [
  // Gmail / Apple Mail — the attribution may wrap across two lines.
  /^\s*On [\s\S]{5,300}?wrote:\s*$/mu,
  // Localized attributions.
  /^\s*Le [\s\S]{5,300}? a écrit\s*:\s*$/mu,
  /^\s*Am [\s\S]{5,300}? schrieb [\s\S]{0,120}?:\s*$/mu,
  /^\s*El [\s\S]{5,300}? escribió:\s*$/mu,
  // Quoted lines.
  /^\s*>/mu,
  // Outlook / desktop clients.
  /^\s*-{2,}\s*Original Message\s*-{2,}\s*$/imu,
  /^\s*-{2,}\s*Forwarded message\s*-{2,}\s*$/imu,
  /^\s*_{5,}\s*$/mu,
  /^\s*From:\s.+\r?\n\s*(?:Sent|Date):\s/mu,
  // Signature separator ("-- ") and the older dashes-only variant.
  /^\s*-{2,}\s*$/mu,
  // Mobile sign-offs.
  /^\s*Sent from my (?:iPhone|iPad|Android|Galaxy|BlackBerry|Windows Phone)/imu,
  /^\s*Sent from (?:Outlook|Mail|Gmail|Yahoo Mail|ProtonMail)\b/imu,
];

export type StripQuotedReplyOptions = {
  /** Hard cap on the returned text (after stripping). */
  maxLength?: number;
};

/** The sender's own words: everything before the first quoted-history or
 *  signature marker, trimmed. Returns an empty string for a reply that was
 *  nothing but quoted history. */
export const stripQuotedReply = (
  text: string | null | undefined,
  options: StripQuotedReplyOptions = {},
) => {
  const source = (text ?? "").replace(/\r\n/gu, "\n");
  let cut = source.length;
  for (const marker of QUOTE_MARKERS) {
    const index = source.search(marker);
    if (index >= 0 && index < cut) cut = index;
  }
  const stripped = source.slice(0, cut).trim();

  return options.maxLength === undefined
    ? stripped
    : stripped.slice(0, options.maxLength);
};

/** A one-line preview of a reply for notifications and list rows. */
export const replyPreview = (text: string, maxLength = 140) => {
  const collapsed = text.replace(/\s+/gu, " ").trim();
  if (collapsed.length <= maxLength) return collapsed;

  return `${collapsed.slice(0, Math.max(0, maxLength - 1)).trimEnd()}…`;
};
