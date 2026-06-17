import { ImapFlow } from "imapflow";
import type { MailboxAuth, NormalizedEmailMessage } from "../types";
import { cleanEmail, directionFor, parseDate } from "../utils";

const DEFAULT_LIMIT = 100;
const DEFAULT_MAILBOX = "INBOX";
const DEFAULT_PORT = 993;

export type ImapMailboxConfig = {
  accountEmail: string;
  auth: MailboxAuth;
  host: string;
  mailbox?: string;
  port?: number;
  secure?: boolean;
};

export type ImapFetchOptions = {
  cursor?: string | null;
  limit?: number;
  since?: Date;
};

export type ImapFetchResult = {
  cursor?: string | null;
  messages: NormalizedEmailMessage[];
};

const addressList = (
  value:
    | {
        address?: string;
        name?: string;
      }[]
    | undefined,
) =>
  (value ?? [])
    .map((item) => ({
      address: cleanEmail(item.address),
      name: item.name ?? null,
    }))
    .filter((item) => item.address);

export const fetchImapMessages = async (
  config: ImapMailboxConfig,
  options: ImapFetchOptions = {},
): Promise<ImapFetchResult> => {
  const client = new ImapFlow({
    auth: config.auth,
    host: config.host,
    logger: false,
    port: config.port ?? DEFAULT_PORT,
    secure: config.secure ?? true,
  });
  await client.connect();
  try {
      const lock = await client.getMailboxLock(config.mailbox ?? DEFAULT_MAILBOX);
    try {
      const mailbox = client.mailbox;
      const exists = mailbox ? mailbox.exists : 0;
      const limit = options.limit ?? DEFAULT_LIMIT;
      const min = Math.max(1, exists - limit + 1);
      const sequence = exists > 0 ? `${min}:*` : "";
      if (!sequence) return { cursor: options.cursor ?? null, messages: [] };
      const messages: NormalizedEmailMessage[] = [];
      let highestUid = Number(options.cursor ?? 0);
      for await (const message of client.fetch(sequence, {
        bodyParts: ["text"],
        envelope: true,
        uid: true,
      })) {
        if (options.cursor && message.uid <= Number(options.cursor)) continue;
        const envelopeDate = message.envelope?.date;
        if (options.since && envelopeDate && envelopeDate < options.since) {
          continue;
        }
        const text = message.bodyParts?.get("text")?.toString("utf8") ?? null;
        const from = addressList(message.envelope?.from)[0] ?? null;
        highestUid = Math.max(highestUid, message.uid);
        messages.push({
          accountEmail: cleanEmail(config.accountEmail),
          bodyText: text,
          direction: directionFor(config.accountEmail, from?.address),
          from,
          id: `imap:${config.accountEmail}:${message.uid}`,
          occurredAt: parseDate(message.envelope?.date?.toISOString()),
          provider: "imap",
          raw: { uid: message.uid },
          snippet: text ? text.slice(0, 240) : null,
          subject: message.envelope?.subject ?? null,
          threadId: message.envelope?.messageId ?? null,
          to: addressList(message.envelope?.to),
        });
      }

      return {
        cursor: highestUid > 0 ? String(highestUid) : options.cursor ?? null,
        messages,
      };
    } finally {
      lock.release();
    }
  } finally {
    await client.logout().catch(() => undefined);
  }
};
