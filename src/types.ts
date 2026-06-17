export type EmailProvider = "gmail" | "microsoft" | "imap" | (string & {});

export type EmailAddress = {
  address: string;
  name?: string | null;
};

export type NormalizedEmailMessage = {
  accountEmail: string;
  bodyText?: string | null;
  direction: "inbound" | "outbound";
  from?: EmailAddress | null;
  id: string;
  occurredAt: Date;
  provider: EmailProvider;
  raw?: unknown;
  snippet?: string | null;
  subject?: string | null;
  threadId?: string | null;
  to: EmailAddress[];
};

export type EmailDeltaResult<TMessage> = {
  cursor?: string | null;
  expired?: boolean;
  messages: TMessage[];
};

export type EmailSubscriptionResult = {
  expiration?: Date | null;
  id?: string | null;
};

export type FetchJsonResult<T> = {
  body: T | null;
  ok: boolean;
  status: number;
};

export type EmailFetch = typeof fetch;

export type TokenCredential = {
  accessToken: string;
};

export type MailboxAuth = {
  pass: string;
  user: string;
};
