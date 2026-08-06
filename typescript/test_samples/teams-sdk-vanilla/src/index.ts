// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// Vanilla Teams SDK sample.
//
// Two things this sample is here to show:
//
//   1. Raw JSON activities. Every command below sends and inspects activities as
//      plain objects instead of typed models, so you can see exactly what goes on
//      the wire and exactly what Azure Bot Service says back.
//   2. Non-Teams channels. teams.ts never checks channelId, so the same handlers
//      run on email. `dump` is the easiest way to see how different a non-Teams
//      payload is.

import { MessageActivity } from '@microsoft/teams.api';
import { App } from '@microsoft/teams.apps';

const APP = new App();

const MENTION_TAG = /<at\b[^>]*>[\s\S]*?<\/at>/gi;

const HELP = [
  'Vanilla teams.ts sample — no Agents SDK, no middleware.',
  '',
  '• help — this message',
  '• dump — reply with the raw JSON of the activity you just sent',
  '• raw — send a message built as a raw JSON object, bypassing the typed models',
  '• custom — try to send a custom activity type and report what ABS answers',
  '• anything else — echo, tagged with the channel it arrived on',
].join('\n');

/** Strip Teams @-mention markup so commands match in channels and group chats. */
function clean(text?: string): string {
  return (text ?? '').replace(MENTION_TAG, '').trim();
}

function block(payload: unknown): string {
  return '```\n' + JSON.stringify(payload, null, 2) + '\n```';
}

/**
 * POST a raw activity object straight to the connector, skipping teams.ts's models.
 *
 * Sending raw means giving up what ActivitySender normally does for you. The one thing
 * it always fills in is `from`, and channels reject the activity without it, so it is
 * added here. Returns the payload as actually sent, plus the connector's answer.
 *
 * ctx.api is rebuilt per activity from the inbound serviceUrl, so this stays correct
 * on email as well as Teams.
 */
async function postRaw(
  ctx: any,
  payload: Record<string, any>
): Promise<[Record<string, any>, number, string]> {
  const bot = ctx.activity.recipient;
  payload.from ??= { id: bot?.id, name: bot?.name };

  const url = `${ctx.api.serviceUrl}/v3/conversations/${ctx.activity.conversation.id}/activities`;
  try {
    const response = await ctx.api.http.post(url, payload);
    return [payload, response.status, JSON.stringify(response.data ?? '')];
  } catch (error: any) {
    // Surfaced to the user rather than swallowed.
    const response = error?.response;
    const status = response?.status ?? 0;
    const body = response?.data ? JSON.stringify(response.data) : String(error?.message ?? error);
    return [payload, status, body];
  }
}

APP.on('message', async (ctx) => {
  const command = clean(ctx.activity.text).toLowerCase();

  if (command === 'help') {
    await ctx.send(HELP);
    return;
  }

  if (command === 'dump') {
    // The typed model round-tripped back to the JSON the channel actually sent.
    await ctx.send(
      `Raw inbound activity (${ctx.activity.channelId}):\n${block(ctx.activity)}`
    );
    return;
  }

  if (command === 'raw') {
    const [sent, status, body] = await postRaw(ctx, {
      type: 'message',
      text: 'Sent as a raw JSON object — no typed model.',
    });
    await ctx.send(
      `Sent ${block(sent)}\nConnector answered **${status}** ${body || '(empty body)'}`
    );
    return;
  }

  if (command === 'custom') {
    const [sent, status, body] = await postRaw(ctx, {
      type: 'vanilla/customActivity',
      text: 'custom activity type',
    });
    await ctx.send(
      `Tried to send a custom activity type: ${block(sent)}\n` +
      `Connector answered **${status}**: ${body || '(empty body)'}`
    );
    return;
  }

  await ctx.send(
    new MessageActivity(`(${ctx.activity.channelId}) You said: ${clean(ctx.activity.text)}`)
  );
});

APP.event('error', ({ error }) => {
  console.error('Unhandled error:', error);
});

APP.start(Number(process.env.PORT ?? 3979)).catch((err) => {
  console.error('failed to start', err);
  process.exit(1);
});
