// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// Live sample: AgentApplication (Agents-for-js) with teams.ts App mounted via
// TeamsSdkMiddleware.
//
// Behavior split:
//   * TEAMS_APP — teams.ts handlers (typed activities, rich builders, ApiClient).
//   * AGENT_SDK_APP — fallthrough handlers, cross-channel commands, and auth.

import { ActivityTypes, Channels } from '@microsoft/agents-activity';
import {
  AgentApplication,
  CloudAdapter,
  loadAuthConfigFromEnv,
  MemoryStorage,
  RouteRank,
  TurnContext,
  type TurnState,
} from '@microsoft/agents-hosting';
import { startServer } from '@microsoft/agents-hosting-express';
import { Client as ApiClient, MessageActivity } from '@microsoft/teams.api';

import { isTeamsChannel, useTeamsSdk } from 'teams-sdk-middleware';

import { helpCard, taskFormCard, taskLauncherCard } from './cards';

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

/**
 * Anchor a command to the *first line* of the message, tolerating a leading
 * @-mention. Email bodies carry a signature and quoted thread below the
 * command, so matching against the whole body would never hit.
 */
function command(name: string): RegExp {
  const mention = '(?:<at\\b[^>]*>[\\s\\S]*?</at>|@\\S+)';
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(`^\\s*(?:${mention}\\s*)*${escaped}[ \\t]*(?:\\r?\\n[\\s\\S]*)?$`, 'i');
}

// ───────────────────────────── Bootstrap ─────────────────────────────

const AUTH_CONFIG = loadAuthConfigFromEnv();
const ADAPTER = new CloudAdapter(AUTH_CONFIG);
// The adapter owns the connection manager; teams.ts mints its outbound tokens
// from that same one so both SDKs share a single credential set.
const CONNECTION_MANAGER = ADAPTER.connectionManager;

const STORAGE = new MemoryStorage();

// Auth handlers are configured in .env, not here:
//   AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__<name>__SETTINGS__AZUREBOTOAUTHCONNECTIONNAME
const AUTH_HANDLER_IDS = [
  ...new Set(
    Object.keys(process.env)
      .map((key) => /^AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__(.+?)__/i.exec(key)?.[1])
      .filter((id): id is string => id !== undefined)
  ),
];

const AGENT_SDK_APP = new AgentApplication<TurnState>({
  storage: STORAGE,
  adapter: ADAPTER,
});

AGENT_SDK_APP.onError(async (context, error) => {
  console.error('Unhandled error:', error);
  await context.sendActivity(`⚠️ ${error?.name ?? 'Error'}: ${error?.message ?? error}`);
});

/**
 * Keep `signin/*` invokes on the Agents SDK side.
 *
 * teams.ts registers `signin/tokenExchange`, `signin/verifyState` and
 * `signin/failure` handlers unconditionally when the App is constructed, so its
 * router always matches them — including for a flow AgentApplication started.
 * Those handlers verify against teams.ts's own `defaultConnectionName` (default
 * `"graph"`), so letting them win would strand the Agents SDK flow and 404 when
 * no ABS connection named `graph` exists.
 *
 * Drop this predicate if you want teams.ts to own sign-in instead.
 */
function agentSdkOwnsSignIn(context: TurnContext): boolean {
  return (
    context.activity.type === ActivityTypes.Invoke &&
    (context.activity.name ?? '').toLowerCase().startsWith('signin/')
  );
}

// One call: extracts credentials from CONNECTION_MANAGER, wires teams.ts's
// outbound token callback to it, constructs the App, and installs
// TeamsSdkMiddleware on AGENT_SDK_APP.adapter.
const TEAMS_APP = useTeamsSdk(AGENT_SDK_APP, CONNECTION_MANAGER, {}, agentSdkOwnsSignIn);

// ════════════════════ TEAMS_APP — Teams SDK feature showcase ════════════════════

TEAMS_APP.message(command('help'), async ({ send }) => {
  await send(new MessageActivity().addCard('adaptive', helpCard() as any));
});

TEAMS_APP.message(command('react'), async ({ send, api, activity }) => {
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

TEAMS_APP.message(command('quote'), async ({ reply }) => {
  // Auto-quotes the inbound message.
  await reply('Quoting your message!');
});

TEAMS_APP.message(command('targeted'), async ({ send, activity }) => {
  // Send a targeted (ephemeral) message visible only to the sender.
  const sender = activity.from;
  const targeted = new MessageActivity('👁️ This message is only visible to you.')
    .withRecipient({ id: sender.id, name: sender.name ?? '', role: 'user' }, true);
  await send(targeted);
});

TEAMS_APP.message(command('task'), async ({ send }) => {
  await send(new MessageActivity().addCard('adaptive', taskLauncherCard() as any));
});

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

// ════════════════════ AGENT_SDK_APP — fallthrough + "agents *" commands ════════════════════
// Reached for Teams turns with no matching teams.ts route, and for every non-Teams channel.

AGENT_SDK_APP.onMessage(command('help'), async (context: TurnContext) => {
  await context.sendActivity(
    '[Agent SDK] Commands: help, channel, whoami, mail, signout, ' +
    'agents sdk react, agents sdk proactive.\n' +
    'Teams-only extras (react, quote, targeted, task) need the Teams SDK routes.'
  );
});

AGENT_SDK_APP.onMessage(command('channel'), async (context: TurnContext) => {
  const via = isTeamsChannel(context.activity)
    ? 'Teams turn with no matching teams.ts route → fell through'
    : 'non-Teams channel → passed straight through';
  await context.sendActivity(`[Agent SDK] channelId=${context.activity.channelId} (${via})`);
});

AGENT_SDK_APP.onMessage(command('agents sdk react'), async (context: TurnContext) => {
  // Reach into teams.ts's API client from an Agents SDK handler.
  // TEAMS_APP.api is pinned to the service URL at App construction; for
  // handlers driven by the Agents SDK (whose activities may arrive on a
  // different service URL) build a per-turn ApiClient against the inbound
  // context.activity.serviceUrl while reusing the shared HTTP client.
  if (!isTeamsChannel(context.activity)) {
    await context.sendActivity(
      "[Agent SDK] 'agents sdk react' needs the Teams reactions API; " +
      `channelId=${context.activity.channelId} returns 404 for it.`
    );
    return;
  }
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

AGENT_SDK_APP.onMessage(command('agents sdk proactive'), async (context: TurnContext) => {
  const convId = context.activity.conversation!.id;
  const api = new ApiClient(context.activity.serviceUrl!, TEAMS_APP.api.http);
  const bot = context.activity.recipient;
  const outgoing = new MessageActivity(
    '[Teams SDK] Proactive message triggered from an Agents SDK handler!'
  );
  // Bypassing teams.ts's ActivitySender means nothing populates `from`, and
  // Direct Line rejects the send without it.
  outgoing.from = { id: bot?.id ?? '', name: bot?.name ?? '', role: 'bot' };
  await api.conversations.activities(convId).create(outgoing);
});

// ═══════════════════════════ Authentication ═══════════════════════════
// Auth lives here and not on the teams.ts side: the auth intercept runs inside
// AgentApplication.run, which the middleware only calls when no teams.ts route
// matches. A teams.ts route would run unauthenticated rather than fail.
//
// Both handlers use the same AAD app but different ABS connections, so each holds its own
// token — signing in for one does not satisfy the other.

/** GET a Graph resource with the token cached for `handler`. null on failure. */
async function graphGet(
  context: TurnContext,
  handler: string,
  resource: string
): Promise<any | null> {
  const token = await AGENT_SDK_APP.authorization.getToken(context, handler);
  if (!token?.token) {
    await context.sendActivity(`[Agent SDK] No token for the '${handler}' handler.`);
    return null;
  }

  const response = await fetch(`${GRAPH_BASE_URL}${resource}`, {
    headers: { Authorization: `Bearer ${token.token}` },
  });
  const body: any = await response.json();
  if (!response.ok) {
    const detail = body?.error?.message ?? JSON.stringify(body);
    await context.sendActivity(
      `[Agent SDK] Graph ${resource} returned ${response.status}: ${detail}`
    );
    return null;
  }
  return body;
}

AGENT_SDK_APP.onMessage(
  command('whoami'),
  async (context: TurnContext) => {
    // Sign-in already completed by the time this runs, so getToken reads from cache.
    const me = await graphGet(context, 'graphuser', '/me');
    if (me) {
      await context.sendActivity(
        `[Agent SDK] ${me.displayName} (${me.userPrincipalName})\n` +
        "Handler 'graphuser' — scope User.Read."
      );
    }
  },
  ['graphuser']
);

AGENT_SDK_APP.onMessage(
  command('mail'),
  async (context: TurnContext) => {
    const data = await graphGet(
      context,
      'graphmail',
      '/me/messages?$top=3&$select=subject,receivedDateTime'
    );
    if (data === null) {
      return;
    }
    const messages: any[] = data.value ?? [];
    if (messages.length === 0) {
      await context.sendActivity('[Agent SDK] Mailbox is empty.');
      return;
    }
    const lines = messages.map((m) => `• ${m.subject || '(no subject)'}`).join('\n');
    await context.sendActivity(
      `[Agent SDK] Latest ${messages.length} message(s):\n${lines}\n` +
      "Handler 'graphmail' — scopes User.Read + Mail.Read."
    );
  },
  ['graphmail']
);

AGENT_SDK_APP.onMessage(command('signout'), async (context: TurnContext, state: TurnState) => {
  for (const handler of AUTH_HANDLER_IDS) {
    await AGENT_SDK_APP.authorization.signOut(context, state, handler);
  }
  await context.sendActivity(`[Agent SDK] Signed out of: ${AUTH_HANDLER_IDS.join(', ')}.`);
});

// Every OAuth card is rendered with the same fixed "Sign in" text, so these callbacks are
// the only way to tell which connection a prompt belonged to.
if (AUTH_HANDLER_IDS.length > 0) {
  AGENT_SDK_APP.authorization.onSignInSuccess(async (context, _state, handlerId) => {
    await context.sendActivity(`[Agent SDK] Signed in via '${handlerId}'.`);
  });

  AGENT_SDK_APP.authorization.onSignInFailure(async (context, _state, handlerId) => {
    await context.sendActivity(`[Agent SDK] Sign-in failed for '${handlerId}'.`);
  });
}

// Sign-in cannot complete on email: Azure Bot Service flattens cards into a static image,
// so the button is inert. A started flow would then swallow every later turn before
// routing, leaving the mailbox silent. These routes outrank the ones above and decline, so
// no flow ever begins.
//
// signout is declined here too. Tokens are keyed by (channelId, userId, connectionName),
// and the email identity (an SMTP address) never holds one, so signing out would always
// report success for zero work — and could never reach a token held on Teams anyway.
const NO_AUTH_CHANNELS: string[] = [Channels.Email];

function blockedAuthSelector(name: string) {
  const pattern = command(name);
  return async (context: TurnContext): Promise<boolean> =>
    context.activity.type === ActivityTypes.Message &&
    NO_AUTH_CHANNELS.includes(context.activity.channelId ?? '') &&
    pattern.test(context.activity.text ?? '');
}

async function declineAuth(context: TurnContext): Promise<void> {
  await context.sendActivity(
    `[Agent SDK] Sign-in isn't supported on ${context.activity.channelId} — the OAuth ` +
    "card renders as a static image here, so it can't be clicked. Tokens are scoped per " +
    'channel, so there is nothing to sign in or out of on this one. ' +
    'Try whoami / mail on Teams or Web Chat.'
  );
}

for (const name of ['whoami', 'mail', 'signout']) {
  AGENT_SDK_APP.addRoute(blockedAuthSelector(name), declineAuth, false, RouteRank.First);
}

// ═══════════════════════════ Fallthrough ═══════════════════════════

AGENT_SDK_APP.onActivity(
  ActivityTypes.Message,
  async (context: TurnContext) => {
    // Default echo fallthrough. Fires when no teams.ts route matches and none
    // of the commands above matched either. Email bodies carry signatures and
    // quoted threads, so echo only the first line.
    let text = (context.activity.text ?? '').trim();
    const firstLine = text.split(/\r?\n/).map((l) => l.trim()).find((l) => l.length > 0) ?? '';
    if (firstLine !== text) {
      text = `${firstLine} […]`;
    }
    await context.sendActivity(`[Agent SDK] (${context.activity.channelId}) You said: ${text}`);
  },
  undefined,
  RouteRank.Last
);

// ─────────────────────────── HTTP wiring ───────────────────────────

const server = startServer(AGENT_SDK_APP);
server.get('/', (_req: any, res: any) => {
  res.send('teams-sdk-sample is running. POST activities to /api/messages.');
});
