export type MicrosoftGraphWebhookBody = {
  value?: {
    clientState?: string;
    subscriptionId?: string;
  }[];
};

export type MicrosoftGraphWebhookPayload = {
  clientState?: string;
  subscriptionId: string;
};

export const parseMicrosoftGraphWebhook = (
  body: MicrosoftGraphWebhookBody,
  expectedClientState?: string | null,
) =>
  (body.value ?? []).filter(
    (item): item is MicrosoftGraphWebhookPayload =>
      Boolean(item.subscriptionId) &&
      (!expectedClientState || item.clientState === expectedClientState),
  );
