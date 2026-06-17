# @absolutejs/email

Provider-neutral email sync adapters for AbsoluteJS applications.

This package owns the provider mechanics:

- Gmail REST sync, `users.watch`, and Pub/Sub payload parsing
- Microsoft Graph message delta sync and subscription notification parsing
- IMAP over TLS for custom mailbox providers
- Normalized message shapes that host apps can persist, score, filter, or enrich

Host applications own product policy:

- user consent and privacy controls
- credential storage
- business-specific relevance scoring
- database persistence
- queue scheduling and UI

## Install

```bash
bun add @absolutejs/email
```

## Gmail

```ts
import {
  createGmailClient,
  gmailMessagesToNormalized,
  parseGmailPubSubWebhook,
} from "@absolutejs/email";

const client = createGmailClient({ accessToken });
const { messages, cursor } = await client.listHistory({ cursor: historyId });
const normalized = await gmailMessagesToNormalized(client, messages, {
  accountEmail: "member@example.com",
});
```

## Microsoft Graph

```ts
import {
  createMicrosoftGraphEmailClient,
  microsoftMessagesToNormalized,
} from "@absolutejs/email";

const client = createMicrosoftGraphEmailClient({ accessToken });
const { messages, cursor } = await client.listDelta();
const normalized = microsoftMessagesToNormalized(messages, {
  accountEmail: "member@example.com",
});
```

## IMAP

```ts
import { fetchImapMessages } from "@absolutejs/email";

const result = await fetchImapMessages({
  accountEmail: "member@example.com",
  auth: { pass: appPassword, user: "member@example.com" },
  host: "imap.fastmail.com",
  port: 993,
  secure: true,
});
```
