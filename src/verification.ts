import type { GmailClient } from "./adapters/gmail";
import { gmailMessagesToNormalized } from "./adapters/gmail";
import type { MicrosoftEmailClient } from "./adapters/microsoft";
import { microsoftMessagesToNormalized } from "./adapters/microsoft";
import type { EmailProvider, NormalizedEmailMessage } from "./types";
import { cleanEmail } from "./utils";

const DEFAULT_CODE_LENGTH = 6;
const DEFAULT_MAX_BODY_BYTES = 128 * 1024;
const DEFAULT_MAX_CANDIDATES = 20;
const DEFAULT_MAX_MARKER_GAP = 16;
const MAX_AUTHENTICATION_HEADER_BYTES = 4_096;
const MAX_AUTHENTICATION_HEADERS = 10;
const MAX_BODY_BYTES = 1024 * 1024;
const MAX_CANDIDATES = 100;
const MAX_LOOKUP_WINDOW_MS = 10 * 60_000;
const MAX_PROFILE_VALUES = 20;
const MAX_PROFILE_VALUE_LENGTH = 256;

export type EmailSenderAuthenticationPolicy = {
  readonly allowedHeaderFromDomains: readonly string[];
  readonly trustedAuthservIds: readonly string[];
};

export type EmailVerificationProfile = {
  readonly bodyMarkers: readonly string[];
  readonly codeLength?: number;
  readonly id: string;
  readonly maxMarkerGap?: number;
  readonly origins: readonly string[];
  readonly providers: readonly EmailProvider[];
  readonly senderAddresses: readonly string[];
  readonly senderAuthentication: EmailSenderAuthenticationPolicy;
  readonly subjectIncludesAny: readonly string[];
};

export type EmailVerificationLookupInput = {
  readonly accountEmail: string;
  readonly maxCandidates?: number;
  readonly notAfter: Date;
  readonly notBefore: Date;
  readonly profile: EmailVerificationProfile;
  readonly requiredBodyText?: readonly string[];
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
    readonly senderAuthenticated: true;
  };
};

export type EmailVerificationErrorCode =
  | "ambiguous_match"
  | "candidate_limit"
  | "invalid_profile"
  | "lookup_failed"
  | "no_match";

const SAFE_ERROR_MESSAGES: Record<EmailVerificationErrorCode, string> = {
  ambiguous_match: "More than one verification message or code matched.",
  candidate_limit: "The verification candidate limit was exceeded.",
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

const invalidProfile = (): never => {
  throw new EmailVerificationError("invalid_profile");
};

const validOrigin = (value: string) => {
  try {
    const parsed = new URL(value);
    return (
      parsed.protocol === "https:" &&
      parsed.origin === value &&
      parsed.username === "" &&
      parsed.password === ""
    );
  } catch {
    return false;
  }
};

const canonicalDomain = (value: string) => {
  if (
    value.length === 0 ||
    value.length > 253 ||
    !/^[A-Za-z0-9.-]+$/u.test(value) ||
    value.startsWith(".") ||
    value.endsWith(".") ||
    value.includes("..")
  ) {
    return "";
  }
  try {
    const hostname = new URL(`https://${value}`).hostname.toLowerCase();
    return hostname === value.toLowerCase() ? hostname : "";
  } catch {
    return "";
  }
};

const validEmail = (value: string) => {
  if (value.length === 0 || value.length > 254 || /[^\x21-\x7e]/u.test(value)) {
    return false;
  }
  const separator = value.lastIndexOf("@");
  if (separator < 1 || separator !== value.indexOf("@")) return false;
  const local = value.slice(0, separator);
  return (
    local.length <= 64 &&
    !local.startsWith(".") &&
    !local.endsWith(".") &&
    !local.includes("..") &&
    /^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+$/u.test(local) &&
    canonicalDomain(value.slice(separator + 1)) !== ""
  );
};

const validValues = (
  values: readonly string[],
  validate: (value: string) => boolean = (value) => value.trim().length > 0,
) =>
  values.length > 0 &&
  values.length <= MAX_PROFILE_VALUES &&
  values.every(
    (value) =>
      value.length <= MAX_PROFILE_VALUE_LENGTH &&
      !/[\u0000-\u001f\u007f]/u.test(value) &&
      validate(value),
  );

const normalizedProfile = (profile: EmailVerificationProfile) => {
  if (
    typeof profile !== "object" ||
    profile === null ||
    typeof profile.id !== "string" ||
    !Array.isArray(profile.bodyMarkers) ||
    !Array.isArray(profile.origins) ||
    !Array.isArray(profile.providers) ||
    !Array.isArray(profile.senderAddresses) ||
    !Array.isArray(profile.subjectIncludesAny) ||
    typeof profile.senderAuthentication !== "object" ||
    profile.senderAuthentication === null ||
    !Array.isArray(profile.senderAuthentication.allowedHeaderFromDomains) ||
    !Array.isArray(profile.senderAuthentication.trustedAuthservIds)
  ) {
    invalidProfile();
  }
  const codeLength = profile.codeLength ?? DEFAULT_CODE_LENGTH;
  const maxMarkerGap = profile.maxMarkerGap ?? DEFAULT_MAX_MARKER_GAP;
  const senders = profile.senderAddresses.map(cleanEmail);
  const allowedDomains =
    profile.senderAuthentication.allowedHeaderFromDomains.map(canonicalDomain);
  const trustedAuthservIds =
    profile.senderAuthentication.trustedAuthservIds.map((value) =>
      value.toLowerCase(),
    );
  if (
    !/^[A-Za-z0-9._-]{1,128}$/u.test(profile.id) ||
    !Number.isSafeInteger(codeLength) ||
    codeLength < 4 ||
    codeLength > 12 ||
    !Number.isSafeInteger(maxMarkerGap) ||
    maxMarkerGap < 0 ||
    maxMarkerGap > 64 ||
    !validValues(profile.bodyMarkers) ||
    !validValues(profile.subjectIncludesAny) ||
    !validValues(profile.origins, validOrigin) ||
    !validValues(profile.providers, (provider) =>
      /^[A-Za-z0-9._-]{1,64}$/u.test(provider),
    ) ||
    !validValues(senders, validEmail) ||
    !validValues(allowedDomains, (domain) => domain !== "") ||
    !validValues(trustedAuthservIds, (value) =>
      /^[A-Za-z0-9.-]{1,253}$/u.test(value),
    ) ||
    new Set(profile.origins).size !== profile.origins.length ||
    new Set(profile.providers).size !== profile.providers.length ||
    new Set(senders).size !== senders.length ||
    new Set(allowedDomains).size !== allowedDomains.length ||
    new Set(trustedAuthservIds).size !== trustedAuthservIds.length ||
    new Set(profile.bodyMarkers.map((value) => value.toLowerCase())).size !==
      profile.bodyMarkers.length ||
    new Set(profile.subjectIncludesAny.map((value) => value.toLowerCase()))
      .size !== profile.subjectIncludesAny.length ||
    senders.some((sender) => !allowedDomains.includes(senderDomain(sender)))
  ) {
    invalidProfile();
  }

  return {
    allowedDomains,
    codeLength,
    maxMarkerGap,
    senders,
    trustedAuthservIds,
  };
};

const validatedInput = (
  input: EmailVerificationLookupInput & {
    readonly expectedOrigin: string;
    readonly maxBodyBytes?: number;
  },
) => {
  const profile = normalizedProfile(input.profile);
  if (
    typeof input.accountEmail !== "string" ||
    !(input.notBefore instanceof Date) ||
    !(input.notAfter instanceof Date) ||
    (input.requiredBodyText !== undefined &&
      !Array.isArray(input.requiredBodyText))
  ) {
    invalidProfile();
  }
  const accountEmail = cleanEmail(input.accountEmail);
  const maxBodyBytes = input.maxBodyBytes ?? DEFAULT_MAX_BODY_BYTES;
  const maxCandidates = input.maxCandidates ?? DEFAULT_MAX_CANDIDATES;
  const notBefore = input.notBefore.getTime();
  const notAfter = input.notAfter.getTime();
  const requiredBodyText = input.requiredBodyText ?? [];
  if (
    !validEmail(accountEmail) ||
    !input.profile.origins.includes(input.expectedOrigin) ||
    !Number.isSafeInteger(maxBodyBytes) ||
    maxBodyBytes < 1 ||
    maxBodyBytes > MAX_BODY_BYTES ||
    !Number.isSafeInteger(maxCandidates) ||
    maxCandidates < 1 ||
    maxCandidates > MAX_CANDIDATES ||
    !Number.isFinite(notBefore) ||
    !Number.isFinite(notAfter) ||
    notAfter < notBefore ||
    notAfter - notBefore > MAX_LOOKUP_WINDOW_MS ||
    (requiredBodyText.length > 0 && !validValues(requiredBodyText))
  ) {
    invalidProfile();
  }
  return {
    ...profile,
    accountEmail,
    maxBodyBytes,
    maxCandidates,
    notAfter,
    notBefore,
    requiredBodyText,
  };
};

const codeOccurrencesAfterMarkers = (
  body: string,
  profile: EmailVerificationProfile,
  codeLength: number,
  maxMarkerGap: number,
) => {
  const lowerBody = body.toLowerCase();
  const codePattern = new RegExp(
    `^[\\s:：#-]{0,${maxMarkerGap}}(\\d{${codeLength}})(?!\\d)`,
    "u",
  );
  const codes: string[] = [];
  for (const configuredMarker of profile.bodyMarkers) {
    const marker = configuredMarker.toLowerCase();
    let start = 0;
    while (start < lowerBody.length) {
      const markerIndex = lowerBody.indexOf(marker, start);
      if (markerIndex < 0) break;
      const valueStart = markerIndex + marker.length;
      const window = body.slice(
        valueStart,
        valueStart + maxMarkerGap + codeLength + 1,
      );
      const match = codePattern.exec(window);
      if (match?.[1] !== undefined) codes.push(match[1]);
      start = Math.max(valueStart, markerIndex + 1);
    }
  }
  return codes;
};

const senderDomain = (address: string) => {
  const separator = address.lastIndexOf("@");
  return separator < 0 ? "" : canonicalDomain(address.slice(separator + 1));
};

const authenticatedSender = (
  message: NormalizedEmailMessage,
  allowedDomains: readonly string[],
  trustedAuthservIds: readonly string[],
) => {
  const headers = message.authenticationResults ?? [];
  if (
    headers.length === 0 ||
    headers.length > MAX_AUTHENTICATION_HEADERS ||
    headers.some((value) => value.length > MAX_AUTHENTICATION_HEADER_BYTES)
  ) {
    return false;
  }
  const trusted = headers.flatMap((value) => {
    const separator = value.indexOf(";");
    if (separator < 1) return [];
    const authservId = value
      .slice(0, separator)
      .trim()
      .split(/\s+/u)[0]
      ?.toLowerCase();
    if (authservId === undefined || !trustedAuthservIds.includes(authservId)) {
      return [];
    }
    const dmarc =
      /(?:^|;)\s*dmarc(?:\/\d+)?\s*=\s*pass\b[^;]{0,1024}\bheader\.from\s*=\s*([A-Za-z0-9.-]{1,253})\b/iu.exec(
        value,
      );
    const domain = canonicalDomain(dmarc?.[1] ?? "");
    return domain === "" ? [] : [domain];
  });
  const from = cleanEmail(message.from?.address);
  const domain = senderDomain(from);
  return (
    trusted.length === 1 &&
    trusted[0] === domain &&
    allowedDomains.includes(domain)
  );
};

const messageMatches = (
  message: NormalizedEmailMessage,
  input: EmailVerificationLookupInput,
  validation: ReturnType<typeof validatedInput>,
) => {
  const occurredAt = message.occurredAt.getTime();
  return (
    message.direction === "inbound" &&
    cleanEmail(message.accountEmail) === validation.accountEmail &&
    input.profile.providers.includes(message.provider) &&
    validation.senders.includes(cleanEmail(message.from?.address)) &&
    Number.isFinite(occurredAt) &&
    occurredAt >= validation.notBefore &&
    occurredAt <= validation.notAfter &&
    input.profile.subjectIncludesAny.some((subject) =>
      (message.subject ?? "").toLowerCase().includes(subject.toLowerCase()),
    ) &&
    authenticatedSender(
      message,
      validation.allowedDomains,
      validation.trustedAuthservIds,
    )
  );
};

export const resolveEmailVerificationCode = (
  messages: readonly NormalizedEmailMessage[],
  input: EmailVerificationLookupInput & {
    readonly expectedOrigin: string;
    readonly maxBodyBytes?: number;
  },
): EmailVerificationCodeResult => {
  const validation = validatedInput(input);
  if (messages.length > validation.maxCandidates) {
    throw new EmailVerificationError("candidate_limit");
  }

  const matches: Array<{ code: string; message: NormalizedEmailMessage }> = [];
  for (const message of messages) {
    if (!messageMatches(message, input, validation)) continue;
    const body = message.bodyText ?? "";
    if (
      body.length > validation.maxBodyBytes ||
      new TextEncoder().encode(body).byteLength > validation.maxBodyBytes ||
      !validation.requiredBodyText.every((value) => body.includes(value))
    ) {
      continue;
    }
    const codes = codeOccurrencesAfterMarkers(
      body,
      input.profile,
      validation.codeLength,
      validation.maxMarkerGap,
    );
    if (codes.length > 1) {
      throw new EmailVerificationError("ambiguous_match");
    }
    if (codes[0] !== undefined) matches.push({ code: codes[0], message });
  }

  if (matches.length === 0) throw new EmailVerificationError("no_match");
  if (matches.length !== 1) {
    throw new EmailVerificationError("ambiguous_match");
  }
  const match = matches[0]!;
  if (
    match.message.id.length === 0 ||
    match.message.id.length > 512 ||
    /[\u0000-\u001f\u007f]/u.test(match.message.id)
  ) {
    throw new EmailVerificationError("no_match");
  }
  return Object.freeze({
    bytes: new TextEncoder().encode(match.code),
    evidence: Object.freeze({
      matchedAt: match.message.occurredAt.getTime(),
      messageId: match.message.id,
      parserId: input.profile.id,
      provider: match.message.provider,
      senderAuthenticated: true,
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
  validatedInput(input);
  let messages: readonly NormalizedEmailMessage[];
  try {
    messages = await lookup.find(input);
  } catch (error) {
    if (error instanceof EmailVerificationError) throw error;
    throw new EmailVerificationError("lookup_failed");
  }
  return resolveEmailVerificationCode(messages, input);
};

const candidateLimit = (
  configured: number | undefined,
  requested: number | undefined,
) => {
  const value = requested ?? configured ?? DEFAULT_MAX_CANDIDATES;
  if (!Number.isSafeInteger(value) || value < 1 || value > MAX_CANDIDATES) {
    invalidProfile();
  }
  return value;
};

export const createGmailVerificationMessageLookup = (input: {
  readonly accountEmail: string;
  readonly client: GmailClient;
  readonly maxCandidates?: number;
}): EmailVerificationMessageLookup => ({
  find: async (query) => {
    const validation = validatedInput({
      ...query,
      expectedOrigin: query.profile.origins[0] ?? "",
    });
    const maxCandidates = candidateLimit(
      input.maxCandidates,
      query.maxCandidates,
    );
    const references = new Map<string, { id: string }>();
    for (const sender of validation.senders) {
      const found = await input.client.searchMessages({
        maxResults: maxCandidates + 1,
        query: `from:${sender} after:${Math.floor(validation.notBefore / 1000)} before:${Math.ceil(validation.notAfter / 1000)}`,
      });
      for (const message of found) references.set(message.id, message);
      if (references.size > maxCandidates) {
        throw new EmailVerificationError("candidate_limit");
      }
    }
    const normalized = await gmailMessagesToNormalized(
      input.client,
      [...references.values()],
      {
        accountEmail: input.accountEmail,
      },
    );
    if (normalized.length !== references.size) {
      throw new EmailVerificationError("lookup_failed");
    }
    return normalized;
  },
});

export const createMicrosoftVerificationMessageLookup = (input: {
  readonly accountEmail: string;
  readonly client: MicrosoftEmailClient;
  readonly maxCandidates?: number;
}): EmailVerificationMessageLookup => ({
  find: async (query) => {
    const validation = validatedInput({
      ...query,
      expectedOrigin: query.profile.origins[0] ?? "",
    });
    const maxCandidates = candidateLimit(
      input.maxCandidates,
      query.maxCandidates,
    );
    const messages = new Map<
      string,
      Awaited<ReturnType<typeof input.client.searchMessages>>[number]
    >();
    for (const sender of validation.senders) {
      for (const message of await input.client.searchMessages({
        maxResults: maxCandidates + 1,
        query: `from:${sender}`,
      })) {
        if (message.id !== undefined) messages.set(message.id, message);
      }
      if (messages.size > maxCandidates) {
        throw new EmailVerificationError("candidate_limit");
      }
    }
    const full = await Promise.all(
      [...messages.keys()].map((id) => input.client.getMessage(id)),
    );
    if (full.some((message) => message === null)) {
      throw new EmailVerificationError("lookup_failed");
    }
    return microsoftMessagesToNormalized(
      full.filter((message) => message !== null),
      { accountEmail: input.accountEmail },
    );
  },
});
