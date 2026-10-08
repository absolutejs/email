import type { Change } from '@absolutejs/changelog';

export const change: Change = {
	detail:
		'The Gmail push validation fix shipped as 0.4.1 on the older 0.4 line, so 0.6 and 0.7 never had it. Gmail Pub/Sub push payloads are validated, and numeric history cursors are normalized only when they are safe integers, so a large cursor is never rounded.',
	kind: 'fixed',
	summary:
		'Validate Gmail push payloads and keep large history cursors exact (the 0.4.1 fix, now on the current line)',
	symbols: ['parseGmailPubSubWebhook']
};
