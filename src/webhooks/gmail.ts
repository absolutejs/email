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

export const parseGmailPubSubWebhook = (
  body: GmailPubSubWebhookBody,
): GmailPubSubPayload | null => {
  const data = body.message?.data;
  if (!data) return null;

  try {
    const decoded = Buffer.from(data, "base64").toString("utf8");
    const parsed = JSON.parse(decoded) as GmailPubSubPayload;

    return parsed.emailAddress ? parsed : null;
  } catch {
    return null;
  }
};
