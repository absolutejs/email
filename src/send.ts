/** Host-bound Gmail sending. The caller owns approvals, credentials and durable
 * execution. There are deliberately no retries: Gmail has no send idempotency key. */
export type OutgoingEmail = {
  from: string;
  to: string[];
  subject: string;
  text: string;
  messageId: string;
  inReplyTo?: string;
  threadId?: string;
  attachments?: { name: string; contentType: string; bytes: Uint8Array }[];
};
const address = (value: string) => {
  if (!/^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(value) || value.length > 254)
    throw new Error('Invalid email address');
  return value;
};
const messageId = (value: string) => {
  if (!/^<[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+>$/.test(value) || value.length > 254)
    throw new Error('Invalid email message identity');
  return value;
};
const encodedHeader = (value: string) => {
  if (/[\r\n\x00-\x1f\x7f]/.test(value)) throw new Error('Invalid email header');
  // Fold between UTF-8 code points; each encoded word stays within RFC 2047's limit.
  const pieces: string[] = []; let piece = '';
  for (const char of value) {
    if (Buffer.byteLength(piece + char) > 42) { pieces.push(piece); piece = ''; }
    piece += char;
  }
  if (piece) pieces.push(piece);
  return pieces.map(part => `=?UTF-8?B?${Buffer.from(part).toString('base64')}?=`).join('\r\n ');
};
const base64Lines = (bytes: Uint8Array | string) => Buffer.from(bytes).toString('base64').match(/.{1,76}/g)?.join('\r\n') ?? '';
export const buildOutgoingMime = (input: OutgoingEmail): string => {
  if (!Array.isArray(input.to) || input.to.length < 1 || input.to.length > 20 || !input.subject.trim() || input.subject.length > 300 || !input.text.trim() || Buffer.byteLength(input.text) > 100000)
    throw new Error('Invalid outgoing email');
  const attachments = input.attachments ?? [];
  if (attachments.length > 5 || attachments.reduce((n, a) => n + a.bytes.byteLength, 0) > 10 * 1024 * 1024)
    throw new Error('Email attachments exceed limit');
  const headers = [`From: ${address(input.from)}`, `To: ${input.to.map(address).join(', ')}`, `Subject: ${encodedHeader(input.subject)}`, `Message-ID: ${messageId(input.messageId)}`, 'MIME-Version: 1.0'];
  if (input.inReplyTo) headers.push(`In-Reply-To: ${messageId(input.inReplyTo)}`, `References: ${messageId(input.inReplyTo)}`);
  if (input.threadId && (!/^[a-zA-Z0-9_-]{1,200}$/.test(input.threadId) || !input.inReplyTo)) throw new Error('Replies require thread and message identities');
  const textPart = ['Content-Type: text/plain; charset=UTF-8', 'Content-Transfer-Encoding: base64', '', base64Lines(input.text)].join('\r\n');
  if (!attachments.length) return [...headers, textPart].join('\r\n');
  const boundary = `absolute_${crypto.randomUUID().replaceAll('-', '')}`;
  const parts = attachments.map(a => {
    if (!/^[a-zA-Z0-9.+-]+\/[a-zA-Z0-9.+-]+$/.test(a.contentType) || !a.name || a.name.length > 200 || /[\r\n\x00-\x1f\x7f]/.test(a.name)) throw new Error('Invalid email attachment');
    const name = encodeURIComponent(a.name).replaceAll("'", '%27');
    return [`Content-Type: ${a.contentType}`, `Content-Disposition: attachment; filename*=UTF-8''${name}`, 'Content-Transfer-Encoding: base64', '', base64Lines(a.bytes)].join('\r\n');
  });
  return [...headers, `Content-Type: multipart/mixed; boundary="${boundary}"`, '', `--${boundary}`, textPart, ...parts.flatMap(p => [`--${boundary}`, p]), `--${boundary}--`, ''].join('\r\n');
};
export class EmailSendError extends Error {
  constructor(public readonly outcome: 'rejected' | 'unknown', public readonly status?: number) {
    super(outcome === 'rejected' ? 'Email provider rejected the send' : 'Email send outcome requires reconciliation');
  }
}
export const sendGmailMessage = async (input: OutgoingEmail, options: {
  fetch: (url: string, init: RequestInit) => Promise<Response>;
  signal?: AbortSignal;
}): Promise<{ id: string; threadId: string; messageId: string }> => {
  const raw = Buffer.from(buildOutgoingMime(input)).toString('base64url');
  let response: Response;
  try {
    response = await options.fetch('https://gmail.googleapis.com/gmail/v1/users/me/messages/send', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ raw, ...(input.threadId ? { threadId: input.threadId } : {}) }), ...(options.signal ? { signal: options.signal } : {})
    });
  } catch { throw new EmailSendError('unknown'); }
  if (!response.ok) throw new EmailSendError([400,401,403,404,413,422,429].includes(response.status) ? 'rejected' : 'unknown', response.status);
  try {
    const result = await response.json() as { id?: unknown; threadId?: unknown };
    if (typeof result.id !== 'string' || !result.id || typeof result.threadId !== 'string' || !result.threadId) throw new Error();
    return { id: result.id, threadId: result.threadId, messageId: input.messageId };
  } catch { throw new EmailSendError('unknown', response.status); }
};
