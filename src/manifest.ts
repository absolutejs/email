import { defineManifest, toolFactory } from "@absolutejs/manifest";
import { Type } from "@sinclair/typebox";
import { fetchImapMessages } from "./adapters/imap";
import type { ImapMailboxConfig } from "./adapters/imap";
import type { GmailClient } from "./adapters/gmail";
import {
  gmailMessagesToNormalized,
  gmailMessageToNormalized,
} from "./adapters/gmail";
import type { MicrosoftEmailClient } from "./adapters/microsoft";
import { microsoftMessagesToNormalized } from "./adapters/microsoft";
import type { EmailSyncStateStore } from "./stores/inMemoryStateStore";
import type { NormalizedEmailMessage } from "./types";

/**
 * Composite runtime (v1 convention: several cooperating pieces, no single
 * instance). The host binds whichever provider clients it constructed for the
 * mailbox it wants AI tools on. Tokens live inside the clients and are never
 * echoed by any tool.
 */
export type EmailManifestRuntime = {
  /** The mailbox the bound clients read. */
  accountEmail: string;
  gmail?: GmailClient;
  microsoft?: MicrosoftEmailClient;
  imap?: ImapMailboxConfig;
  /** Per-account sync cursors + webhook subscription state. */
  stateStore?: EmailSyncStateStore;
};

const tool = toolFactory<EmailManifestRuntime>();

const MAX_LIST_LIMIT = 50;
const DEFAULT_LIST_LIMIT = 10;
const SNIPPET_LENGTH = 240;
const DETAIL_BODY_LENGTH = 4000;

/* Tools return a compact projection: never `raw` (provider payloads can carry
 * headers/ids the host didn't mean to expose) and never credentials. */
const summarize = (message: NormalizedEmailMessage) => ({
  direction: message.direction,
  from: message.from?.address ?? null,
  id: message.id,
  occurredAt: message.occurredAt.toISOString(),
  snippet:
    message.snippet ?? message.bodyText?.slice(0, SNIPPET_LENGTH) ?? null,
  subject: message.subject ?? null,
  threadId: message.threadId ?? null,
  to: message.to.map((recipient) => recipient.address),
});

const listRecent = async (runtime: EmailManifestRuntime, limit: number) => {
  if (runtime.microsoft) {
    const state = await runtime.stateStore?.get(
      "microsoft",
      runtime.accountEmail,
    );
    const delta = await runtime.microsoft.listDelta({
      cursor: state?.cursor ?? null,
      pageSize: limit,
    });

    return microsoftMessagesToNormalized(delta.messages.slice(0, limit), {
      accountEmail: runtime.accountEmail,
    });
  }
  if (runtime.imap) {
    const result = await fetchImapMessages(runtime.imap, { limit });

    return result.messages.slice(-limit);
  }
  if (runtime.gmail) {
    const state = await runtime.stateStore?.get("gmail", runtime.accountEmail);
    if (!state?.cursor) {
      return "gmail listing needs a stored history cursor — run a watch/history sync for this mailbox first (see the gmail-sync wiring recipe)";
    }
    const delta = await runtime.gmail.listHistory({ cursor: state.cursor });

    return gmailMessagesToNormalized(
      runtime.gmail,
      delta.messages.slice(-limit),
      {
        accountEmail: runtime.accountEmail,
      },
    );
  }

  return "no email client is bound for this mailbox";
};

/* @absolutejs/email is provider mechanics (Gmail / Microsoft Graph / IMAP
 * clients, normalized messages, webhook parsers) — every entry point is a
 * function taking host-owned credentials, so there is no serializable
 * top-level config: settings are empty and each provider is a wiring recipe.
 * OAuth access tokens are per-user material minted by the host's auth layer,
 * NOT env keys; only the IMAP mailbox and the at-rest encryption key are
 * env-shaped. */
export const manifest = defineManifest<
  Record<never, never>,
  EmailManifestRuntime
>()({
  contract: 2,
  identity: {
    accent: "#ea4335",
    category: "messaging",
    description:
      "Provider-neutral email sync: Gmail REST history + `users.watch` Pub/Sub, Microsoft Graph delta + change subscriptions, and IMAP over TLS — all normalized to one message shape your app can persist, score, or enrich. The host owns consent, credential storage, and persistence; this package owns the provider mechanics.",
    docsUrl: "https://github.com/absolutejs/email",
    name: "@absolutejs/email",
    tagline: "Read and sync your users’ email — Gmail, Outlook, and IMAP.",
  },
  requires: {
    env: [
      {
        description:
          "Key used to encrypt stored mailbox credentials at rest (any long random string). Only needed when you persist tokens with encryptEmailSecret.",
        key: "EMAIL_ENCRYPTION_KEY",
        optional: true,
        secret: true,
      },
      {
        description:
          "IMAP server hostname — only for the imap-sync recipe (custom mailbox providers).",
        example: "imap.fastmail.com",
        key: "IMAP_HOST",
        optional: true,
      },
      {
        description:
          "IMAP account (the mailbox email address) — only for the imap-sync recipe.",
        example: "member@yoursite.com",
        key: "IMAP_USER",
        optional: true,
      },
      {
        description:
          "IMAP app password for the account — only for the imap-sync recipe.",
        key: "IMAP_PASSWORD",
        optional: true,
        secret: true,
      },
    ],
  },
  settings: Type.Object({}),
  tools: {
    list_recent_messages: tool.runtime({
      annotations: { idempotentHint: true, openWorldHint: true },
      authorization: {
        approval: "never",
        audience: "owner",
        destinations: ["configured-mailbox-provider"],
        effects: ["read", "external-network"],
        idempotency: { mode: "host" },
        requiredScopes: ["email:read"],
        reversible: false,
      },
      description:
        "List recent messages in the connected mailbox (sender, recipients, subject, snippet — never full provider payloads or credentials). Uses whichever provider client the host bound: Microsoft Graph delta, IMAP recent, or Gmail history from the stored cursor.",
      handler: async ({ limit }, runtime) => {
        const result = await listRecent(runtime, limit ?? DEFAULT_LIST_LIMIT);

        return typeof result === "string"
          ? result
          : JSON.stringify(result.map(summarize));
      },
      input: Type.Object({
        limit: Type.Optional(
          Type.Integer({
            default: DEFAULT_LIST_LIMIT,
            maximum: MAX_LIST_LIMIT,
            minimum: 1,
          }),
        ),
      }),
    }),
    message_detail: tool.runtime({
      annotations: { idempotentHint: true, openWorldHint: true },
      authorization: {
        approval: "never",
        audience: "owner",
        destinations: ["configured-mailbox-provider"],
        effects: ["read", "external-network"],
        idempotency: { mode: "host" },
        requiredScopes: ["email:read"],
        resource: { idField: "id", type: "email-message" },
        reversible: false,
      },
      description:
        "Fetch one message by id, with its text body (truncated). Supported for Gmail mailboxes; Microsoft Graph and IMAP only expose messages through the list sync.",
      handler: async ({ id }, runtime) => {
        if (!runtime.gmail) {
          return "message_detail is only available for Gmail mailboxes — use list_recent_messages for this provider";
        }
        const message = await runtime.gmail.getMessage(id);
        const normalized = message
          ? gmailMessageToNormalized(message, {
              accountEmail: runtime.accountEmail,
            })
          : null;
        if (!normalized) return `no message found with id "${id}"`;

        return JSON.stringify({
          ...summarize(normalized),
          bodyText: normalized.bodyText?.slice(0, DETAIL_BODY_LENGTH) ?? null,
        });
      },
      input: Type.Object({ id: Type.String({ minLength: 1 }) }),
    }),
    sync_status: tool.runtime({
      annotations: { readOnlyHint: true },
      authorization: {
        approval: "never",
        audience: "owner",
        effects: ["read"],
        requiredScopes: ["email:read"],
      },
      description:
        "Report sync state for the connected mailbox: which providers are bound, whether each has a stored cursor, and when the change-notification subscription expires. Never returns cursors or credentials.",
      handler: async (_input, runtime) => {
        const providers = (
          [
            ["gmail", runtime.gmail],
            ["microsoft", runtime.microsoft],
            ["imap", runtime.imap],
          ] as const
        ).filter(([, client]) => client !== undefined);
        if (providers.length === 0) return "no email client is bound";
        const statuses = await Promise.all(
          providers.map(async ([provider]) => {
            const state =
              (await runtime.stateStore?.get(provider, runtime.accountEmail)) ??
              null;

            return {
              hasCursor: Boolean(state?.cursor),
              hasSubscription: Boolean(state?.subscriptionId),
              provider,
              subscriptionExpiresAt:
                state?.subscriptionExpiration?.toISOString() ?? null,
            };
          }),
        );

        return JSON.stringify({
          accountEmail: runtime.accountEmail,
          providers: statuses,
        });
      },
      input: Type.Object({}),
    }),
  },
  wiring: [
    {
      description:
        "Sync a member's Gmail: pull changed messages since the stored history cursor and normalize them. Pair with users.watch + parseGmailPubSubWebhook for push notifications.",
      id: "default",
      server: {
        code: [
          "// The host owns OAuth: mint the member's Gmail access token (scope",
          "// https://www.googleapis.com/auth/gmail.readonly) from your auth layer.",
          "// TODO: const accessToken = ...; const accountEmail = ...;",
          "const gmail = createGmailClient({ accessToken });",
          "const emailSyncState = createInMemoryEmailSyncStateStore();",
          "",
          "// On each sync (or Pub/Sub webhook): pull history and normalize.",
          "// const state = await emailSyncState.get('gmail', accountEmail);",
          "// const delta = await gmail.listHistory({ cursor: state?.cursor });",
          "// const messages = await gmailMessagesToNormalized(gmail, delta.messages, { accountEmail });",
          "// await emailSyncState.set({ accountEmail, cursor: delta.cursor, provider: 'gmail' });",
        ].join("\n"),
        imports: [
          {
            from: "@absolutejs/email",
            names: [
              "createGmailClient",
              "createInMemoryEmailSyncStateStore",
              "gmailMessagesToNormalized",
            ],
          },
        ],
        placement: "module-scope",
      },
      title: "Gmail sync (OAuth access token from your auth layer)",
    },
    {
      description:
        "Sync an Outlook/Microsoft 365 mailbox with Graph message delta. Pair with createOrRenewSubscription + parseMicrosoftGraphWebhook for change notifications.",
      id: "microsoft-sync",
      server: {
        code: [
          "// The host owns OAuth: mint the member's Microsoft Graph access token",
          "// (Mail.Read) from your auth layer.",
          "// TODO: const accessToken = ...; const accountEmail = ...;",
          "const outlook = createMicrosoftGraphEmailClient({ accessToken });",
          "",
          "// On each sync (or Graph webhook): pull the delta and normalize.",
          "// const delta = await outlook.listDelta({ cursor: storedCursor });",
          "// const messages = microsoftMessagesToNormalized(delta.messages, { accountEmail });",
        ].join("\n"),
        imports: [
          {
            from: "@absolutejs/email",
            names: [
              "createMicrosoftGraphEmailClient",
              "microsoftMessagesToNormalized",
            ],
          },
        ],
        placement: "module-scope",
      },
      title: "Outlook / Microsoft 365 sync (Graph delta)",
    },
    {
      description:
        "Fetch a custom provider's mailbox over IMAP with an app password. The returned cursor (highest UID) makes the next fetch incremental.",
      id: "imap-sync",
      server: {
        code: [
          "const imapResult = await fetchImapMessages({",
          "\taccountEmail: ${env.IMAP_USER} ?? '',",
          "\tauth: { pass: ${env.IMAP_PASSWORD} ?? '', user: ${env.IMAP_USER} ?? '' },",
          "\thost: ${env.IMAP_HOST} ?? ''",
          "});",
          "// imapResult.messages are normalized; persist imapResult.cursor for the next fetch.",
        ].join("\n"),
        imports: [{ from: "@absolutejs/email", names: ["fetchImapMessages"] }],
        placement: "module-scope",
      },
      title: "IMAP mailbox (app password)",
    },
  ],
});
