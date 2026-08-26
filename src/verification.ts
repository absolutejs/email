import type { GmailClient } from "./adapters/gmail";
import { gmailMessagesToNormalized } from "./adapters/gmail";
import type { MicrosoftEmailClient } from "./adapters/microsoft";
import { microsoftMessagesToNormalized } from "./adapters/microsoft";
import type { EmailProvider, NormalizedEmailMessage } from "./types";
import { cleanEmail } from "./utils";

const DEFAULT_CODE_LENGTH = 6;
const DEFAULT_MAX_BODY_BYTES = 128 * 1024;
const DEFAULT_MAX_CANDIDATES = 20;
const DEFAULT_MAX_MARKER_GAP = 32;

export type EmailVerificationProfile = {
  readonly bodyMarkers: readonly string[];
  readonly codeLength?: number;
  readonly id: string;
  readonly maxMarkerGap?: number;
  readonly origins: readonly string[];
  readonly providers: readonly EmailProvider[];
  readonly senderAddresses: readonly string[];
  readonly subjectIncludesAny?: readonly string[];
};

export type EmailVerificationLookupInput = {
  readonly accountEmail: string;
  readonly notAfter: Date;
  readonly notBefore: Date;
  readonly profile: EmailVerificationProfile;
};

export type EmailVerificationMessageLookup = {
  readonly find: (
    input: EmailVerificationLookupInput,
  ) => Promise<readonly NormalizedEmailMessage[]>;
};

export type EmailVerificationCodeResult = {
  readonly bytes: Uint8Array;
  readonly evidence: {
    readonly matchedAt: number;
    readonly messageId: string;
    readonly parserId: string;
    readonly provider: string;
  };
};

export type EmailVerificationErrorCode =
  "ambiguous_match" | "invalid_profile" | "lookup_failed" | "no_match";

const SAFE_ERROR_MESSAGES: Record<EmailVerificationErrorCode, string> = {
  ambiguous_match: "More than one verification message or code matched.",
  invalid_profile: "The email verification profile is invalid.",
  lookup_failed: "The mailbox lookup failed safely.",
  no_match: "No verification message matched the exact profile.",
};

export class EmailVerificationError extends Error {
  readonly code: EmailVerificationErrorCode;

  constructor(code: EmailVerificationErrorCode) {
    super(SAFE_ERROR_MESSAGES[code]);
    this.code = code;
    this.name = "EmailVerificationError";
  }
}

const validOrigin = (value: string) => {
  try {
    const parsed = new URL(value);
    return parsed.protocol === "https:" && parsed.origin === value;
  } catch {
    return false;
  }
};

const normalizedProfile = (profile: EmailVerificationProfile) => {
  const codeLength = profile.codeLength ?? DEFAULT_CODE_LENGTH;
  const maxMarkerGap = profile.maxMarkerGap ?? DEFAULT_MAX_MARKER_GAP;
  const senders = profile.senderAddresses.map(cleanEmail);
  if (
    profile.id.trim().length === 0 ||
    !Number.isSafeInteger(codeLength) ||
    codeLength < 4 ||
    codeLength > 12 ||
    !Number.isSafeInteger(maxMarkerGap) ||
    maxMarkerGap < 0 ||
    maxMarkerGap > 128 ||
    profile.bodyMarkers.length === 0 ||
    profile.bodyMarkers.some((marker) => marker.trim().length === 0) ||
    profile.origins.length === 0 ||
    profile.origins.some((origin) => !validOrigin(origin)) ||
    profile.providers.length === 0 ||
    senders.length === 0 ||
    senders.some((sender) => sender.length === 0)
  ) {
    throw new EmailVerificationError("invalid_profile");
  }

  return { codeLength, maxMarkerGap, senders };
};

const codesAfterMarkers = (
  body: string,
  profile: EmailVerificationProfile,
  codeLength: number,
  maxMarkerGap: number,
) => {
  const lowerBody = body.toLocaleLowerCase("en-US");
  const codePattern = new RegExp(`(?:^|\\D)(\\d{${codeLength}})(?!\\d)`, "u");
  const codes = new Set<string>();
  for (const configuredMarker of profile.bodyMarkers) {
    const marker = configuredMarker.toLocaleLowerCase("en-US");
    let start = 0;
    while (start < lowerBody.length) {
      const markerIndex = lowerBody.indexOf(marker, start);
      if (markerIndex < 0) break;
      const valueStart = markerIndex + marker.length;
      const window = body.slice(
        valueStart,
        valueStart + maxMarkerGap + codeLength + 2,
      );
      const match = codePattern.exec(window);
      if (match?.[1] !== undefined) codes.add(match[1]);
      start = valueStart;
    }
  }
  return codes;
};

const messageMatches = (
  message: NormalizedEmailMessage,
  input: EmailVerificationLookupInput,
  senders: readonly string[],
) => {
  const occurredAt = message.occurredAt.getTime();
  const subjects = input.profile.subjectIncludesAny ?? [];
  return (
    message.direction === "inbound" &&
    cleanEmail(message.accountEmail) === cleanEmail(input.accountEmail) &&
    input.profile.providers.includes(message.provider) &&
    senders.includes(cleanEmail(message.from?.address)) &&
    Number.isFinite(occurredAt) &&
    occurredAt >= input.notBefore.getTime() &&
    occurredAt <= input.notAfter.getTime() &&
    (subjects.length === 0 ||
      subjects.some((subject) =>
        (message.subject ?? "")
          .toLocaleLowerCase("en-US")
          .includes(subject.toLocaleLowerCase("en-US")),
      ))
  );
};

export const resolveEmailVerificationCode = (
  messages: readonly NormalizedEmailMessage[],
  input: EmailVerificationLookupInput & {
    readonly expectedOrigin: string;
    readonly maxBodyBytes?: number;
  },
): EmailVerificationCodeResult => {
  const { codeLength, maxMarkerGap, senders } = normalizedProfile(
    input.profile,
  );
  if (!input.profile.origins.includes(input.expectedOrigin)) {
    throw new EmailVerificationError("invalid_profile");
  }
  const maxBodyBytes = input.maxBodyBytes ?? DEFAULT_MAX_BODY_BYTES;
  if (!Number.isSafeInteger(maxBodyBytes) || maxBodyBytes < 1) {
    throw new EmailVerificationError("invalid_profile");
  }

  const matches: Array<{ code: string; message: NormalizedEmailMessage }> = [];
  for (const message of messages) {
    if (!messageMatches(message, input, senders)) continue;
    const body = message.bodyText ?? "";
    if (new TextEncoder().encode(body).byteLength > maxBodyBytes) continue;
    const codes = codesAfterMarkers(
      body,
      input.profile,
      codeLength,
      maxMarkerGap,
    );
    if (codes.size > 1) {
      throw new EmailVerificationError("ambiguous_match");
    }
    const [code] = codes;
    if (code !== undefined) matches.push({ code, message });
  }

  if (matches.length === 0) throw new EmailVerificationError("no_match");
  if (matches.length !== 1) {
    throw new EmailVerificationError("ambiguous_match");
  }
  const match = matches[0]!;
  return Object.freeze({
    bytes: new TextEncoder().encode(match.code),
    evidence: Object.freeze({
      matchedAt: match.message.occurredAt.getTime(),
      messageId: match.message.id,
      parserId: input.profile.id,
      provider: match.message.provider,
    }),
  });
};

export const retrieveEmailVerificationCode = async (
  lookup: EmailVerificationMessageLookup,
  input: EmailVerificationLookupInput & {
    readonly expectedOrigin: string;
    readonly maxBodyBytes?: number;
  },
) => {
  let messages: readonly NormalizedEmailMessage[];
  try {
    messages = await lookup.find(input);
  } catch {
    throw new EmailVerificationError("lookup_failed");
  }
  return resolveEmailVerificationCode(messages, input);
};

export const createGmailVerificationMessageLookup = (input: {
  readonly accountEmail: string;
  readonly client: GmailClient;
  readonly maxCandidates?: number;
}): EmailVerificationMessageLookup => ({
  find: async (query) => {
    const maxCandidates = input.maxCandidates ?? DEFAULT_MAX_CANDIDATES;
    const references = new Map<string, { id: string }>();
    for (const sender of query.profile.senderAddresses) {
      const found = await input.client.searchMessages({
        maxResults: maxCandidates,
        query: `from:${cleanEmail(sender)} after:${Math.floor(query.notBefore.getTime() / 1000)} before:${Math.ceil(query.notAfter.getTime() / 1000)}`,
      });
      for (const message of found) references.set(message.id, message);
    }
    return gmailMessagesToNormalized(
      input.client,
      [...references.values()].slice(0, maxCandidates),
      { accountEmail: input.accountEmail },
    );
  },
});

export const createMicrosoftVerificationMessageLookup = (input: {
  readonly accountEmail: string;
  readonly client: MicrosoftEmailClient;
  readonly maxCandidates?: number;
}): EmailVerificationMessageLookup => ({
  find: async (query) => {
    const maxCandidates = input.maxCandidates ?? DEFAULT_MAX_CANDIDATES;
    const messages = new Map<
      string,
      Awaited<ReturnType<typeof input.client.searchMessages>>[number]
    >();
    for (const sender of query.profile.senderAddresses) {
      for (const message of await input.client.searchMessages({
        maxResults: maxCandidates,
        query: `from:${cleanEmail(sender)}`,
      })) {
        if (message.id !== undefined) messages.set(message.id, message);
      }
    }
    return microsoftMessagesToNormalized(
      [...messages.values()].slice(0, maxCandidates),
      { accountEmail: input.accountEmail },
    );
  },
});
