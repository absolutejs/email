import type { GmailClient, GmailMessage } from "./adapters/gmail";

export const EMAIL_PLACEMENTS = [
  "inbox",
  "spam",
  "trash",
  "other",
  "missing",
] as const;

export type EmailPlacement = (typeof EMAIL_PLACEMENTS)[number];

export type EmailPlacementResult = {
  messageId: string | null;
  placement: EmailPlacement;
};

export const gmailPlacementFromLabels = (
  labelIds: GmailMessage["labelIds"],
): Exclude<EmailPlacement, "missing"> => {
  const labels = new Set(labelIds ?? []);
  if (labels.has("SPAM")) return "spam";
  if (labels.has("TRASH")) return "trash";
  if (labels.has("INBOX")) return "inbox";

  return "other";
};

export const findGmailMessagePlacement = async (
  client: GmailClient,
  input: { query: string },
): Promise<EmailPlacementResult> => {
  const [match] = await client.searchMessages({
    maxResults: 1,
    query: input.query,
  });
  if (!match) return { messageId: null, placement: "missing" };
  const message = await client.getMessage(match.id);
  if (!message) return { messageId: match.id, placement: "missing" };

  return {
    messageId: message.id ?? match.id,
    placement: gmailPlacementFromLabels(message.labelIds),
  };
};
