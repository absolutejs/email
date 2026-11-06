import { expect, test } from 'bun:test';
import { buildOutgoingMime, sendGmailMessage } from './send';
const input = { from: 'sender@example.test', to: ['recipient@example.test'], subject: 'Hello 🌎', text: 'A real\nmessage', messageId: '<test@example.test>' };
test('MIME encodes UTF-8 and rejects header injection before transport', async () => {
  const mime = buildOutgoingMime(input);
  expect(mime).toContain('Subject: =?UTF-8?B?');
  expect(mime).toContain(Buffer.from(input.text).toString('base64'));
  for (const changed of [{ subject: 'bad\r\nBcc: victim@example.test' }, { from: 'x@example.test\r\nX: x' }, { inReplyTo: '<x>\r\nX: x' }])
    expect(() => buildOutgoingMime({ ...input, ...changed })).toThrow();
  expect(() => buildOutgoingMime({ ...input, threadId: 'abc' })).toThrow();
});
test('Gmail send is single-attempt and distinguishes rejection from unknown', async () => {
  let calls = 0;
  await expect(sendGmailMessage(input, { fetch: async () => { calls++; throw new Error('private transport failure'); } })).rejects.toMatchObject({ outcome: 'unknown' });
  expect(calls).toBe(1);
  await expect(sendGmailMessage(input, { fetch: async () => new Response('', {status:403}) })).rejects.toMatchObject({ outcome: 'rejected', status:403 });
  await expect(sendGmailMessage(input, { fetch: async () => new Response('{}') })).rejects.toMatchObject({ outcome: 'unknown' });
  expect(await sendGmailMessage(input, { fetch: async (_url, init) => {
    const wire = JSON.parse(String(init.body)); expect(Buffer.from(wire.raw, 'base64url').toString()).toContain('Message-ID: <test@example.test>');
    return Response.json({id:'sent',threadId:'thread'});
  } })).toEqual({id:'sent',threadId:'thread',messageId:input.messageId});
});
test('attachments use bounded multipart MIME', () => {
  const mime = buildOutgoingMime({...input,attachments:[{name:'résumé.docx',contentType:'application/octet-stream',bytes:new Uint8Array([1,2,3])}]});
  expect(mime).toContain('multipart/mixed'); expect(mime).toContain('AQID'); expect(mime).toContain("filename*=UTF-8''r%C3%A9sum%C3%A9.docx");
  expect(() => buildOutgoingMime({...input,attachments:[{name:'x',contentType:'bad\r\nX: y',bytes:new Uint8Array()}]})).toThrow();
});
