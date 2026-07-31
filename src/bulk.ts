// Whether a message was sent by a machine is answered by headers that only an
// adapter ever sees — the normalized message hands consumers `raw` and leaves
// each of them to re-discover which provider puts headers where, and which of
// the dozen conventions actually mean "bulk".
//
// These signals are not equal, and collapsing them into one boolean is what
// makes downstream trust rules impossible to write. A `List-Unsubscribe` is
// the sender's own infrastructure declaring itself a mailing system —
// authoritative, and true even for a sender in your address book. An
// `Auto-Submitted` is an RFC 3834 auto-reply. `Precedence: bulk` is a weaker,
// older convention. A consumer deciding "should this bypass my contact-list
// exemption?" needs to know WHICH fired, so they are reported separately.

export type HeaderLike = { name?: string | null; value?: string | null };

export type BulkMailSignals = {
  /** RFC 3834 `Auto-Submitted` other than "no" — a machine-generated reply. */
  autoSubmitted: boolean;
  /** `Precedence: bulk|junk|list`, or Microsoft's suppression hint. */
  bulkPrecedence: boolean;
  /** `List-Id` / `List-Unsubscribe` / `Feedback-ID` — the sender declares
   *  itself a mailing system. The strongest signal here: a sender does not
   *  attach an unsubscribe link to a personal message. */
  mailingList: boolean;
};

const NONE: BulkMailSignals = {
  autoSubmitted: false,
  bulkPrecedence: false,
  mailingList: false,
};

const BULK_PRECEDENCE = new Set(["bulk", "junk", "list"]);

const readHeader = (
  headers: HeaderLike[] | Record<string, string | null | undefined>,
  name: string,
) => {
  if (Array.isArray(headers)) {
    const found = headers.find(
      (header) => (header.name ?? "").toLowerCase() === name,
    );

    return found?.value ?? null;
  }
  const entry = Object.entries(headers).find(
    ([key]) => key.toLowerCase() === name,
  );

  return entry?.[1] ?? null;
};

/** Classify a message's headers into the separate machine-sender signals.
 *  Accepts either the array shape Gmail/Graph return or a plain header map. */
export const bulkSignalsFromHeaders = (
  headers:
    | HeaderLike[]
    | Record<string, string | null | undefined>
    | null
    | undefined,
): BulkMailSignals => {
  if (!headers) return NONE;
  const autoSubmitted = readHeader(headers, "auto-submitted")?.toLowerCase();
  const precedence = readHeader(headers, "precedence")?.toLowerCase();

  return {
    autoSubmitted: Boolean(autoSubmitted) && autoSubmitted !== "no",
    bulkPrecedence:
      (precedence !== null &&
        precedence !== undefined &&
        BULK_PRECEDENCE.has(precedence)) ||
      Boolean(readHeader(headers, "x-auto-response-suppress")),
    mailingList: Boolean(
      readHeader(headers, "list-id") ??
        readHeader(headers, "list-unsubscribe") ??
        readHeader(headers, "feedback-id"),
    ),
  };
};

/** Whether any machine-sender signal fired. Convenience for consumers that
 *  genuinely only need a yes/no — prefer the individual signals when the
 *  answer feeds a trust decision. */
export const isBulkMail = (signals: BulkMailSignals) =>
  signals.autoSubmitted || signals.bulkPrecedence || signals.mailingList;
