# @absolutejs/email

The Gmail adapter can also locate a tagged delivery canary across Inbox, Spam,
Trash, and other labels with `findGmailMessagePlacement`. The host application
owns canary scheduling, mailbox credentials, alert policy, and persistence.

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

## Tool-confined verification codes

`@absolutejs/email/verification` provides deterministic Gmail, Microsoft Graph,
and IMAP lookups plus a strict parser for Agent Exchange source adapters. It is a
library API, deliberately absent from the package's model-facing manifest tools.

The main verification entry point is browser-safe for Gmail and Microsoft Graph.
Import the server-only IMAP lookup from `@absolutejs/email/verification/imap` so
browser bundles never pull in Node TLS or `imapflow`.

Profiles bind a mailbox lookup to exact HTTPS origins, provider names, sender
addresses, subject markers, body markers, a short time window, and one code
length. Retrieval fails closed if there is no match, more than one matching
message, or more than one marker-bound code.

```ts
import {
  createGmailVerificationMessageLookup,
  retrieveEmailVerificationCode,
} from "@absolutejs/email/verification";

const result = await retrieveEmailVerificationCode(
  createGmailVerificationMessageLookup({ accountEmail, client: gmail }),
  {
    accountEmail,
    expectedOrigin: "https://accounts.example.com",
    notBefore: new Date(requestCreatedAt),
    notAfter: new Date(requestExpiresAt),
    profile: {
      bodyMarkers: ["verification code"],
      id: "accounts-example-six-digit-v1",
      origins: ["https://accounts.example.com"],
      providers: ["gmail"],
      senderAddresses: ["security@example.com"],
      subjectIncludesAny: ["sign in"],
    },
  },
);

// Pass result.bytes directly to a trusted sink or E2EE envelope. Do not stringify
// it, log it, add it to a prompt, or return it from an agent tool.
result.bytes.fill(0);
```

Email codes remain bearer values and are not phishing-resistant. Prefer OAuth,
passkeys, or provider-native delegated actions whenever the upstream service
supports them.

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
