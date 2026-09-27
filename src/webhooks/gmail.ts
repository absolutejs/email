export type GmailPubSubWebhookBody = {
  message?: {
    data?: string;
    messageId?: string;
  };
  subscription?: string;
};

export type GmailPubSubPayload = {
  emailAddress?: string;
  historyId?: string;
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

export const parseGmailPubSubWebhook = (
  body: unknown,
): GmailPubSubPayload | null => {
  if (!isRecord(body) || !isRecord(body.message)) return null;
  const data = body.message.data;
  if (typeof data !== "string" || !data) return null;

  try {
    const decoded = Buffer.from(data, "base64").toString("utf8");
    const parsed: unknown = JSON.parse(decoded);
    if (
      !isRecord(parsed) ||
      typeof parsed.emailAddress !== "string" ||
      !parsed.emailAddress.trim()
    )
      return null;

    const payload: GmailPubSubPayload = { emailAddress: parsed.emailAddress };
    const cursor = parsed.historyId;
    if (typeof cursor === "string" && /^\d+$/u.test(cursor)) {
      payload.historyId = cursor;
    } else if (
      typeof cursor === "number" &&
      Number.isSafeInteger(cursor) &&
      cursor >= 0
    ) {
      payload.historyId = String(cursor);
    }
    // Never stringify an unsafe number: JSON parsing may already have rounded
    // it. Keep the notification so consumers can sync from their stored cursor.
    return payload;
  } catch {
    return null;
  }
};
