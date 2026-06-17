export type EmailSyncState = {
  accountEmail: string;
  cursor?: string | null;
  provider: string;
  subscriptionExpiration?: Date | null;
  subscriptionId?: string | null;
};

export type EmailSyncStateStore = {
  get: (provider: string, accountEmail: string) => Promise<EmailSyncState | null>;
  set: (state: EmailSyncState) => Promise<void>;
};

export const createInMemoryEmailSyncStateStore = (): EmailSyncStateStore => {
  const states = new Map<string, EmailSyncState>();
  const key = (provider: string, accountEmail: string) =>
    `${provider}:${accountEmail.toLowerCase()}`;

  return {
    get: async (provider, accountEmail) =>
      states.get(key(provider, accountEmail)) ?? null,
    set: async (state) => {
      states.set(key(state.provider, state.accountEmail), state);
    },
  };
};
