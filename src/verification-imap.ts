import type { ImapFetchOptions, ImapMailboxConfig } from "./adapters/imap";
import { fetchImapMessages } from "./adapters/imap";
import type { EmailVerificationMessageLookup } from "./verification";

const DEFAULT_MAX_CANDIDATES = 20;

export const createImapVerificationMessageLookup = (input: {
  readonly config: ImapMailboxConfig;
  readonly fetch?: (
    config: ImapMailboxConfig,
    options: ImapFetchOptions,
  ) => ReturnType<typeof fetchImapMessages>;
  readonly maxCandidates?: number;
}): EmailVerificationMessageLookup => ({
  find: async (query) => {
    const fetcher = input.fetch ?? fetchImapMessages;
    const result = await fetcher(input.config, {
      limit: input.maxCandidates ?? DEFAULT_MAX_CANDIDATES,
      since: query.notBefore,
    });
    return result.messages;
  },
});
