// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// Live sample: AgentApplication (Agents-for-js) with teams.ts App mounted via
// TeamsSdkMiddleware. Ports the Python sample at
//   Agents-for-python/test_samples/teams_sdk_sample/app.py
//
// Behavior split:
//   * TEAMS_APP — teams.ts handlers (typed activities, rich builders, ApiClient).
//   * AGENT_SDK_APP — fallthrough handlers + cross-channel commands.

import {
  AgentApplication,
  loadAuthConfigFromEnv,
  MemoryStorage,
  MsalConnectionManager,
  TurnContext,
  type TurnState,
} from '@microsoft/agents-hosting';
import { startServer } from '@microsoft/agents-hosting-express';
import { Client as ApiClient, MessageActivity } from '@microsoft/teams.api';

import { agentSdkTurnContext, useTeamsSdk } from 'teams-sdk-middleware';

import {
  helpCard,
  pingCard,
  taskFormCard,
  taskLauncherCard,
} from './cards';

// ───────────────────────────── Bootstrap ─────────────────────────────

const CONNECTION_MANAGER = new MsalConnectionManager(
  undefined,
  undefined,
  loadAuthConfigFromEnv()
);

const AGENT_SDK_APP = new AgentApplication<TurnState>({
  storage: new MemoryStorage(),
});

AGENT_SDK_APP.onError(async (context, error) => {
  console.error('Unhandled error:', error);
  await context.sendActivity(`⚠️ ${error?.name ?? 'Error'}: ${error?.message ?? error}`);
});

// One call: extracts credentials from CONNECTION_MANAGER, wires teams.ts's
// outbound token callback to it, constructs the App, and installs
// TeamsSdkMiddleware on AGENT_SDK_APP.adapter.
const TEAMS_APP = useTeamsSdk(AGENT_SDK_APP, CONNECTION_MANAGER);

// ════════════════════ TEAMS_APP — Teams SDK feature showcase ════════════════════

TEAMS_APP.message('help', async ({ send }) => {
  await send(new MessageActivity().addCard('adaptive', helpCard() as any));
});

TEAMS_APP.message('cards', async ({ send }) => {
  // Send an Adaptive Card. Pressing the button fires Action.Execute, routed to
  // the 'card.action.ping' handler below; teams.ts wraps the typed response in
  // an InvokeResponse and TeamsSdkMiddleware propagates it back through the
  // Agents SDK adapter so the HTTP layer writes the correct response.
  await send(new MessageActivity().addCard('adaptive', pingCard() as any));
});

TEAMS_APP.message('citation', async ({ send }) => {
  const activity = new MessageActivity('Here is a response with citations [1].')
    .addCitation(1, {
      name: 'Teams SDK Documentation',
      abstract: 'Documentation for the Microsoft Teams SDK.',
      url: 'https://microsoft.github.io/teams-sdk/welcome',
      usageInfo: {
        '@id': 'sensitivity-1',
        name: 'Confidential',
        description: 'This information is confidential and for internal use only.',
      },
    } as any)
    .addAiGenerated()
    .addFeedback();
  await send(activity);
});

TEAMS_APP.message('stream', async ({ stream }) => {
  stream.update('Thinking…');
  stream.emit('Streaming is a powerful feature ');
  stream.emit('for sending long responses ');
  stream.emit('incrementally.');
  await stream.close();
});

TEAMS_APP.message('react', async ({ send, api, activity }) => {
  const response = await send("React to this message! I'll add 👍 and remove it.");
  const convId = activity.conversation.id;
  try {
    await api.reactions.add(convId, response.id, 'like');
    await new Promise((r) => setTimeout(r, 2000));
    await api.reactions.delete(convId, response.id, 'like');
  } catch (err) {
    console.error('react: reactions API call failed', err);
  }
});

TEAMS_APP.message('quote', async ({ reply }) => {
  // Auto-quotes the inbound message.
  await reply('Quoting your message!');
});

TEAMS_APP.message('targeted', async ({ send, activity }) => {
  // Send a targeted (ephemeral) message visible only to the sender.
  const sender = activity.from;
  const targeted = new MessageActivity('👁️ This message is only visible to you.')
    .withRecipient({ id: sender.id, name: sender.name ?? '', role: 'user' }, true);
  await send(targeted);
});

TEAMS_APP.message('proactive', async ({ send, activity }) => {
  const convId = activity.conversation.id;
  await send('Proactive message coming in ~3s…');
  setTimeout(() => {
    TEAMS_APP.send(convId, new MessageActivity('📣 Proactive message!')).catch((err) =>
      console.error('proactive: TEAMS_APP.send failed', err)
    );
  }, 3000);
});

TEAMS_APP.message('task', async ({ send }) => {
  await send(new MessageActivity().addCard('adaptive', taskLauncherCard() as any));
});

TEAMS_APP.message('turn context', async ({ send }) => {
  // agentSdkTurnContext() returns the live Agents SDK TurnContext that
  // TeamsSdkMiddleware built for this turn, so this handler can call into the
  // Agents SDK outbound pipeline (and read/write turnState) without spinning
  // up a second context.
  const agentSdkCtx = agentSdkTurnContext();
  await send('[Teams SDK] Sending via teams.ts ActivityContext…');
  await agentSdkCtx.sendActivity(
    '[Agent SDK] Sending via Agents SDK TurnContext from inside a teams.ts handler.'
  );
});

// ─── Adaptive Card invoke handler ─────────────────────────────────

TEAMS_APP.on('card.action.ping' as any, (async ({ activity, send }: any) => {
  const action = activity.value.action;
  console.log(`INVOKE adaptiveCard/action verb=${action.verb} data=${JSON.stringify(action.data)}`);
  await send(`[Teams SDK] Adaptive Card action received. Data: ${JSON.stringify(action.data)}`);
  return { statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: 'Action handled.' };
}) as any);

// ─── Task module (dialog) handlers ────────────────────────────────

TEAMS_APP.on('dialog.open', (async (_ctx: any) => {
  return {
    task: {
      type: 'continue',
      value: {
        title: 'Sample Task Module',
        card: { contentType: 'application/vnd.microsoft.card.adaptive', content: taskFormCard() },
      },
    },
  };
}) as any);

TEAMS_APP.on('dialog.submit', (async ({ activity, send }: any) => {
  const data = activity.value.data;
  await send(`[Teams SDK] Task module submitted. Data: ${JSON.stringify(data)}`);
  return { task: { type: 'message', value: 'Done.' } };
}) as any);

// ─── Other Teams events ───────────────────────────────────────────

TEAMS_APP.on('messageReaction', async ({ activity, send }) => {
  const added = activity.reactionsAdded ?? [];
  const removed = activity.reactionsRemoved ?? [];
  await send(
    `[Teams SDK] Reactions: added=[${added.map((r: any) => r.type).join(',')}] removed=[${removed
      .map((r: any) => r.type)
      .join(',')}]`
  );
});

TEAMS_APP.on('message.submit' as any, (async ({ send }: any) => {
  await send('[Teams SDK] Thanks for your feedback!');
}) as any);

TEAMS_APP.on('meetingStart', async ({ send }) => {
  await send('[Teams SDK] Meeting has started!');
});

TEAMS_APP.on('meetingEnd', async ({ send }) => {
  await send('[Teams SDK] Meeting has ended!');
});

// ════════════════════ AGENT_SDK_APP — fallthrough + "agents *" commands ════════════════════
// These fire for Teams activities that have no matching teams.ts route
// (TeamsSdkMiddleware falls through) and for any non-Teams channel.

AGENT_SDK_APP.onConversationUpdate('membersAdded', async (context: TurnContext) => {
  const added = context.activity.membersAdded ?? [];
  const botId = context.activity.recipient?.id;
  if (added.some((m) => m.id !== botId)) {
    await context.sendActivity('[Agent SDK] Welcome! (This is the Agents SDK welcome path.)');
  }
});

AGENT_SDK_APP.onMessage('agents sdk react', async (context: TurnContext) => {
  // Reach into teams.ts's API client from an Agents SDK handler.
  // TEAMS_APP.api is pinned to the service URL at App construction; for
  // handlers driven by the Agents SDK (whose activities may arrive on a
  // different service URL) build a per-turn ApiClient against the inbound
  // context.activity.serviceUrl while reusing the shared HTTP client.
  const response = await context.sendActivity(
    '[Agent SDK] Adding then removing 👍 via teams.ts API client…'
  );
  const convId = context.activity.conversation!.id;
  const api = new ApiClient(context.activity.serviceUrl!, TEAMS_APP.api.http);
  try {
    await api.reactions.add(convId, response!.id!, 'like');
    await new Promise((r) => setTimeout(r, 2000));
    await api.reactions.delete(convId, response!.id!, 'like');
  } catch (err) {
    console.error('agents sdk react: reactions API call failed', err);
  }
});

AGENT_SDK_APP.onMessage('agents sdk proactive', async (context: TurnContext) => {
  const convId = context.activity.conversation!.id;
  const api = new ApiClient(context.activity.serviceUrl!, TEAMS_APP.api.http);
  await api.conversations
    .activities(convId)
    .create(new MessageActivity('[Teams SDK] Proactive message triggered from an Agents SDK handler!'));
});

AGENT_SDK_APP.onMessage('agents sdk citation', async (context: TurnContext) => {
  // Build a Teams citation activity in Agents SDK code, send via teams.ts.
  // Uses a per-turn ApiClient so the send targets the inbound service URL.
  const convId = context.activity.conversation!.id;
  const teamsMessage = new MessageActivity(
    'Agent SDK handler built this citation using Teams SDK types [1].'
  )
    .addCitation(1, {
      name: 'Agent SDK Documentation',
      abstract: 'Documentation for the Microsoft 365 Agents SDK.',
      url: 'https://learn.microsoft.com/en-us/microsoft-365/agents-sdk/',
      usageInfo: {
        '@id': 'sensitivity-1',
        name: 'Confidential',
        description: 'This information is confidential and for internal use only.',
      },
    } as any)
    .addAiGenerated()
    .addFeedback();
  const api = new ApiClient(context.activity.serviceUrl!, TEAMS_APP.api.http);
  await api.conversations.activities(convId).create(teamsMessage);
});

AGENT_SDK_APP.onActivity('message', async (context: TurnContext) => {
  // Default echo fallthrough. Fires when no teams.ts route matches and none
  // of the "agents *" commands above matched either.
  const text = (context.activity.text ?? '').trim();
  await context.sendActivity(`[Agent SDK] You said: ${text}`);
});

// ─────────────────────────── HTTP wiring ───────────────────────────

const server = startServer(AGENT_SDK_APP);
server.get('/', (_req: any, res: any) => {
  res.send('teams-sdk-sample is running. POST activities to /api/messages.');
});
