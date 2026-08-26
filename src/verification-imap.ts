import type { ImapFetchOptions, ImapMailboxConfig } from "./adapters/imap";
import { fetchImapMessages } from "./adapters/imap";
import type { EmailVerificationMessageLookup } from "./verification";

const DEFAULT_MAX_CANDIDATES = 20;
const MAX_CANDIDATES = 100;

export const createImapVerificationMessageLookup = (input: {
  readonly config: ImapMailboxConfig;
  readonly fetch?: (
    config: ImapMailboxConfig,
    options: ImapFetchOptions,
  ) => ReturnType<typeof fetchImapMessages>;
  readonly maxCandidates?: number;
}): EmailVerificationMessageLookup => ({
  find: async (query) => {
    const maxCandidates =
      query.maxCandidates ?? input.maxCandidates ?? DEFAULT_MAX_CANDIDATES;
    if (
      !Number.isSafeInteger(maxCandidates) ||
      maxCandidates < 1 ||
      maxCandidates > MAX_CANDIDATES
    ) {
      throw new Error("invalid candidate limit");
    }
    const fetcher = input.fetch ?? fetchImapMessages;
    const result = await fetcher(input.config, {
      limit: maxCandidates + 1,
      since: query.notBefore,
    });
    return result.messages;
  },
});
