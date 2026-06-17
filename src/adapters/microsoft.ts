import type {
  EmailDeltaResult,
  EmailFetch,
  EmailSubscriptionResult,
  NormalizedEmailMessage,
  TokenCredential,
} from "../types";
import { cleanEmail, directionFor, fetchJson, parseDate, stripHtml } from "../utils";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const DEFAULT_PAGE_SIZE = 150;

export type MicrosoftGraphRecipient = {
  emailAddress?: {
    address?: string;
    name?: string;
  };
};

export type MicrosoftGraphMessage = {
  body?: { content?: string; contentType?: string };
  bodyPreview?: string | null;
  conversationId?: string;
  from?: MicrosoftGraphRecipient;
  id?: string;
  internetMessageId?: string;
  receivedDateTime?: string;
  sentDateTime?: string;
  subject?: string | null;
  toRecipients?: MicrosoftGraphRecipient[];
};

export type MicrosoftGraphMessagePage = {
  "@odata.deltaLink"?: string;
  "@odata.nextLink"?: string;
  value?: MicrosoftGraphMessage[];
};

export type MicrosoftGraphSubscription = {
  clientState?: string;
  expirationDateTime?: string;
  id?: string;
};

export type MicrosoftEmailClient = {
  createOrRenewSubscription: (input: {
    changeType?: string;
    clientState: string;
    expiration: Date;
    notificationUrl: string;
    resource?: string;
    subscriptionId?: string | null;
  }) => Promise<EmailSubscriptionResult>;
  listDelta: (input?: {
    cursor?: string | null;
    pageSize?: number;
  }) => Promise<EmailDeltaResult<MicrosoftGraphMessage>>;
};

const isPage = (value: unknown): value is MicrosoftGraphMessagePage =>
  typeof value === "object" && value !== null;

const isSubscription = (
  value: unknown,
): value is MicrosoftGraphSubscription =>
  typeof value === "object" && value !== null;

const deltaUrl = (pageSize: number) => {
  const params = new URLSearchParams({
    $select:
      "id,internetMessageId,conversationId,subject,bodyPreview,body,from,toRecipients,receivedDateTime,sentDateTime",
    $top: String(pageSize),
  });

  return `${GRAPH_BASE}/me/messages/delta?${params.toString()}`;
};

const graphEmail = (recipient: MicrosoftGraphRecipient | undefined) =>
  cleanEmail(recipient?.emailAddress?.address);

const graphName = (recipient: MicrosoftGraphRecipient | undefined) =>
  recipient?.emailAddress?.name?.trim() || null;

export const createMicrosoftGraphEmailClient = (
  credential: TokenCredential,
  fetcher: EmailFetch = fetch,
): MicrosoftEmailClient => ({
  createOrRenewSubscription: async (input) => {
    const response = await fetchJson<MicrosoftGraphSubscription>(
      input.subscriptionId
        ? `${GRAPH_BASE}/subscriptions/${input.subscriptionId}`
        : `${GRAPH_BASE}/subscriptions`,
      credential.accessToken,
      {
        body: JSON.stringify({
          changeType: input.changeType ?? "created,updated",
          clientState: input.clientState,
          expirationDateTime: input.expiration.toISOString(),
          notificationUrl: input.notificationUrl,
          resource: input.resource ?? "me/messages",
        }),
        method: input.subscriptionId ? "PATCH" : "POST",
      },
      fetcher,
    );
    if (!response.ok || !isSubscription(response.body)) {
      return { expiration: null, id: input.subscriptionId ?? null };
    }

    return {
      expiration: response.body.expirationDateTime
        ? new Date(response.body.expirationDateTime)
        : input.expiration,
      id: response.body.id ?? input.subscriptionId ?? null,
    };
  },
  listDelta: async (input = {}) => {
    const response = await fetchJson<MicrosoftGraphMessagePage>(
      input.cursor ?? deltaUrl(input.pageSize ?? DEFAULT_PAGE_SIZE),
      credential.accessToken,
      undefined,
      fetcher,
    );
    if (!response.ok || !isPage(response.body)) {
      return { cursor: input.cursor ?? null, messages: [] };
    }

    return {
      cursor:
        response.body["@odata.deltaLink"] ??
        response.body["@odata.nextLink"] ??
        input.cursor ??
        null,
      messages: response.body.value ?? [],
    };
  },
});

export const microsoftMessageToNormalized = (
  message: MicrosoftGraphMessage,
  input: { accountEmail: string },
): NormalizedEmailMessage | null => {
  if (!message.id) return null;
  const fromEmail = graphEmail(message.from);
  const bodyText =
    stripHtml(message.body?.content) || message.bodyPreview || null;

  return {
    accountEmail: cleanEmail(input.accountEmail),
    bodyText,
    direction: directionFor(input.accountEmail, fromEmail),
    from: fromEmail
      ? { address: fromEmail, name: graphName(message.from) }
      : null,
    id: `microsoft:${message.id}`,
    occurredAt: parseDate(
      directionFor(input.accountEmail, fromEmail) === "outbound"
        ? message.sentDateTime
        : message.receivedDateTime,
    ),
    provider: "microsoft",
    raw: message,
    snippet: message.bodyPreview ?? null,
    subject: message.subject ?? null,
    threadId: message.conversationId ?? null,
    to:
      message.toRecipients
        ?.map((recipient) => ({
          address: graphEmail(recipient),
          name: graphName(recipient),
        }))
        .filter((recipient) => recipient.address) ?? [],
  };
};

export const microsoftMessagesToNormalized = (
  messages: MicrosoftGraphMessage[],
  input: { accountEmail: string },
) =>
  messages
    .map((message) => microsoftMessageToNormalized(message, input))
    .filter((message): message is NormalizedEmailMessage => message !== null);
