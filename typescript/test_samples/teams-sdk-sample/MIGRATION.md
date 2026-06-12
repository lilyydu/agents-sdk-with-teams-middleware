# Migration Guide

> **Audience:** Existing developers using `TeamsActivityHandler` (or the older `TeamsAgentExtension`) from `@microsoft/agents-hosting-teams` who want to move to the new Teams SDK middleware (`useTeamsSdk`), which embeds the standalone Teams SDK (`@microsoft/teams.*`) inside an Agents SDK `AgentApplication`.

---

## Introduction

The Agents SDK historically shipped a Teams "extension" in `@microsoft/agents-hosting-teams`. That extension exposed Teams invokes (`task/fetch`, `composeExtension/query`, meeting lifecycle, …) as methods on a `TeamsActivityHandler` wrapper around your `AgentApplication`:

```ts
import { TeamsActivityHandler } from '@microsoft/agents-hosting-teams';

class MyBot extends TeamsActivityHandler {
  protected async handleTeamsTaskModuleFetch(context, request) {
    // ...
  }
}
```

That surface covers Teams *invoke* shapes, but stops there. Things outside of pure invoke handling — building rich activities (citations, AI labels, feedback, streaming, quotes), calling Teams APIs (reactions, conversation members, meeting participants), or sending proactively with full Teams metadata — were either missing, partial, or you had to reach into raw `ConnectorClient`/`TeamsInfo` and assemble JSON yourself.

The **new** approach replaces `TeamsActivityHandler` with a thin **bridge middleware** that lets you embed a full `@microsoft/teams.apps` `App` (the standalone Teams SDK) alongside your `AgentApplication`. You keep the Agents SDK as your hosting front door for non-Teams channels; for Teams turns, the bridge hands the activity to the Teams SDK so you get the full Teams SDK developer experience — typed activities, builders, an `ApiClient` with the full Teams surface, and rich routing — without giving up your existing `AgentApplication` handlers, auth, or storage.

```ts
import { useTeamsSdk } from 'teams-sdk-middleware';

const TEAMS_APP = useTeamsSdk(AGENT_SDK_APP, CONNECTION_MANAGER);

TEAMS_APP.message('task', async (ctx) => {
  // ...
});
```

This document covers what changes, why, and how to migrate handler-by-handler.

---

## Motivation

| | Old `TeamsActivityHandler` | New `useTeamsSdk` bridge |
|---|---|---|
| **Handler surface** | Subset of Teams invokes, hand-curated | Full Teams SDK — every activity, every invoke, every event |
| **Activity model** | Flat `Activity` + `activity.value: any` | Typed discriminated unions (`MessageActivity`, `TaskFetchInvokeActivity`, …) |
| **Outbound builders** | Manual objects + Adaptive Card JSON | Activity builders with first-class citations, AI labels, feedback, streaming |
| **API client** | `ConnectorClient` + `TeamsInfo` helper for a few endpoints | Full `ApiClient` — `.reactions`, `.conversations.activities`, `.meetings`, `.users`, `.teams` |
| **Streaming, citations, quotes** | Not supported / manual | First-class |
| **Release cadence** | Tied to `@microsoft/agents-hosting-teams` | Teams SDK ships independently — new Teams features available the moment teams.ts releases them |
| **Hosting** | Agents SDK owns the HTTP endpoint, auth, storage | Same — Agents SDK still owns hosting; bridge is purely additive |

Net effect: Teams developers get the full Teams SDK feature set; Agents SDK developers keep their existing app, auth, and multi-channel surface. The bridge is ~150 lines of glue.

---

## Basic concepts

### What is the middleware and how it works

`TeamsSdkMiddleware` is a regular Agents SDK [`Middleware`](https://learn.microsoft.com/en-us/azure/bot-service/bot-builder-create-middleware) installed on `AGENT_SDK_APP.adapter`. Every turn passes through it; for **Teams turns** it acts as a router, for **other channels** it is a pure pass-through.

```
inbound HTTP /api/messages
        │
        ▼
CloudAdapter ── AGENT_SDK_APP.adapter.use(TeamsSdkMiddleware)
        │
        ├─ channel != "msteams"  ───────────► AgentApplication handlers (unchanged)
        │
        └─ channel == "msteams"
                ├─ cast Agents SDK Activity → teams.ts Activity (same BF schema)
                ├─ await TEAMS_APP.initialize() (idempotent)
                ├─ expose TurnContext via AsyncLocalStorage
                ├─ TEAMS_APP route matches?
                │     ├─ yes → run teams.ts handler; propagate invoke response
                │     └─ no  → fall through to AgentApplication handlers
```

Three points worth noting:

1. **The middleware never double-dispatches.** If a teams.ts route matches, `AgentApplication`'s own handlers do not also fire for that activity. If no teams.ts route matches, the activity flows to `AgentApplication` exactly as if the bridge weren't installed.
2. **Initialization is per-turn, idempotent.** You don't need an explicit `await TEAMS_APP.initialize()` in your startup code — the middleware does it lazily so that `AgentApplication` handlers can safely call into `TEAMS_APP` even on the first turn.
3. **Invoke responses are propagated through the Agents SDK send pipeline.** Teams SDK invoke handlers return typed responses (e.g. `MessagingExtensionResponse`); the middleware wraps them in an `invokeResponse` activity so the Agents SDK HTTP layer writes the synchronous body. You don't have to think about this.

### Introduction to the Teams SDK and how it's different

The Teams SDK (`@microsoft/teams.api`, `@microsoft/teams.apps`, `@microsoft/teams.cards`, `@microsoft/teams.common`) is the Teams team's standalone TypeScript SDK for building Teams apps. It is **the same SDK you would use without the Agents SDK at all** — the bridge just lets you mount it on top of `AgentApplication`.

Three things to know:

#### Handlers

The Teams SDK uses a small set of fluent registration methods on the `App` object — `app.on(name, cb)` for typed events, `app.message(pattern, cb)` for message-text patterns, and `app.use(cb)` for middleware:

```ts
import { App, ActivityContext } from '@microsoft/teams.apps';

TEAMS_APP.on('message', async (ctx) => { /* any message */ });

TEAMS_APP.message('help', async (ctx) => { /* matches activity.text */ });

TEAMS_APP.on('dialog.open', async (ctx) => { /* TaskFetchInvokeActivity */ });

TEAMS_APP.on('message.ext.query', async (ctx) => { /* compose ext query */ });
```

Compared to `TeamsActivityHandler`:

* No `protected handleTeams*` methods to override — every activity type and invoke shape has a route name like `'dialog.open'` or `'message.ext.query'`.
* Handler signature is `(ctx: ActivityContext<T>) => …` — a single context object — instead of `(context, request)`. The context carries the typed activity, an API client, helpers like `ctx.send`, `ctx.reply`, `ctx.stream`, and access to the inbound activity model.
* You don't unwrap `activity.value` yourself — the typed activity class on `ctx.activity` already has the parsed fields.

See **Handler mapping** below for a one-to-one cheat sheet.

#### Activity model

The Teams SDK models every activity shape as a discriminated union. Where the Agents SDK gives you `{ type: 'invoke', name: 'task/fetch', value: {…} }`, the Teams SDK gives you a `TaskFetchInvokeActivity` with `activity.value.data`, `activity.value.context`, etc. as typed attributes.

```ts
// Agents SDK
async function fetch(context: TurnContext) {
  const commandId = context.activity.value?.data?.commandId;
}

// Teams SDK
TEAMS_APP.on('dialog.open', async (ctx) => {
  const commandId = ctx.activity.value.data.commandId;
});
```

Outbound activities use **builders** that compose features cleanly:

```ts
import { MessageActivity } from '@microsoft/teams.api';

const message = new MessageActivity('See [1] for details.')
  .addCitation(1, { name: 'Docs', abstract: '…', url: '…' })
  .addAiGenerated()
  .addFeedback();

await ctx.send(message);
```

#### Underlying clients

The Teams SDK exposes `@microsoft/teams.api`'s `Client` (commonly imported as `ApiClient`) — a single object with the **full Teams API surface**, grouped logically:

| Client | Old equivalent | What it does |
|---|---|---|
| `api.conversations.activities(convId)` | `ConnectorClient.conversations.sendToConversation` | Send, update, delete, reply to activities |
| `api.conversations.members(convId)` | `TeamsInfo.getTeamMembers` | Roster / membership |
| `api.reactions` | (manual JSON) | Add / remove message reactions |
| `api.meetings` | `TeamsInfo.getMeetingInfo` | Meetings, participants |
| `api.teams` | `TeamsInfo.getTeamDetails` | Team metadata |
| `api.users` | (manual Graph) | User profiles |
| `api.bots` | n/a | Bot self-info |

Inside a teams.ts handler you usually use `ctx.api` (auto-scoped to the inbound `serviceUrl`). Inside an `AgentApplication` handler you build one per-turn — see [Proactive flows](#proactive-flows) below.

---

## Initial setup

Replace the old extension wiring with a single call to `useTeamsSdk` *after* you build your `AgentApplication`:

### Before

```ts
import { AgentApplication, MemoryStorage } from '@microsoft/agents-hosting';
import { TeamsActivityHandler } from '@microsoft/agents-hosting-teams';

class MyBot extends TeamsActivityHandler {
  protected async handleTeamsMessagingExtensionQuery(context, query) {
    /* ... */
  }
  protected async handleTeamsTaskModuleFetch(context, request) {
    /* ... */
  }
}

const AGENT_SDK_APP = new AgentApplication({ storage: new MemoryStorage() });
// Wire the handler to your adapter, etc.
```

### After

```ts
import { AgentApplication, MemoryStorage, MsalConnectionManager } from '@microsoft/agents-hosting';
import { useTeamsSdk } from 'teams-sdk-middleware';

const CONNECTION_MANAGER = new MsalConnectionManager();
const AGENT_SDK_APP = new AgentApplication({ storage: new MemoryStorage() });
const TEAMS_APP = useTeamsSdk(AGENT_SDK_APP, CONNECTION_MANAGER);

TEAMS_APP.on('message.ext.query', async (ctx) => { /* ... */ });
TEAMS_APP.on('dialog.open', async (ctx) => { /* ... */ });
```

`useTeamsSdk` does three things:

1. **Extracts credentials** from `CONNECTION_MANAGER.getDefaultConnectionConfiguration()` — the same `clientId` / `tenantId` your Agents SDK app is already configured with.
2. **Bridges the token provider** — outbound calls from the Teams SDK use the Agents SDK's `MsalConnectionManager`, so you don't double-configure auth.
3. **Installs `TeamsSdkMiddleware`** on `AGENT_SDK_APP.adapter`.

Everything else — your `AgentApplicationOptions`, `CloudAdapter`, storage, error handlers, message handlers, your auth handlers, your Express server — stays exactly as it is. The Agents SDK is still your hosting layer.

---

## Guides

### Handler mapping

teams.ts uses string-keyed routes on `App.on(name, cb)` plus a couple of convenience methods (`app.message`, `app.use`). The following tables map old `TeamsActivityHandler` methods to the corresponding teams.ts route name.

#### Message extensions (`composeExtension/*`)

| Old (`TeamsActivityHandler`) | New (`TEAMS_APP.on(...)`) | Inbound activity |
|---|---|---|
| `handleTeamsMessagingExtensionQuery` | `'message.ext.query'` | `MessageExtensionQueryInvokeActivity` |
| `handleTeamsMessagingExtensionSelectItem` | `'message.ext.select-item'` | `MessageExtensionSelectItemInvokeActivity` |
| `handleTeamsMessagingExtensionSubmitAction` | `'message.ext.submit'` | `MessageExtensionSubmitActionInvokeActivity` |
| `handleTeamsMessagingExtensionBotMessagePreviewEdit` | `'message.ext.submit'` (filter on `ctx.activity.value.botMessagePreviewAction === 'edit'`) | same |
| `handleTeamsMessagingExtensionBotMessagePreviewSend` | `'message.ext.submit'` (filter `... === 'send'`) | same |
| `handleTeamsMessagingExtensionFetchTask` | `'message.ext.open'` | `MessageExtensionFetchTaskInvokeActivity` |
| `handleTeamsAppBasedLinkQuery` | `'message.ext.query-link'` | `MessageExtensionQueryLinkInvokeActivity` |
| `handleTeamsAnonymousAppBasedLinkQuery` | `'message.ext.anon-query-link'` | `MessageExtensionAnonQueryLinkInvokeActivity` |
| `handleTeamsMessagingExtensionConfigurationQuerySettingUrl` | `'message.ext.query-settings-url'` | `MessageExtensionQuerySettingUrlInvokeActivity` |
| `handleTeamsMessagingExtensionConfigurationSetting` | `'message.ext.setting'` | `MessageExtensionSettingInvokeActivity` |
| `handleTeamsMessagingExtensionCardButtonClicked` | `'message.ext.card-button-clicked'` | `MessageExtensionCardButtonClickedInvokeActivity` |

#### Task modules → "dialogs"

The Teams SDK renames *task modules* to **dialogs** (matches the Teams JS SDK 2.x rename).

| Old (`TeamsActivityHandler`) | New (`TEAMS_APP.on(...)`) | Inbound activity |
|---|---|---|
| `handleTeamsTaskModuleFetch` | `'dialog.open'` | `TaskFetchInvokeActivity` |
| `handleTeamsTaskModuleSubmit` | `'dialog.submit'` | `TaskSubmitInvokeActivity` |

The response shapes also changed — see the [Task module response shape](#task-module-response-shape-gotcha) note below.

#### Adaptive cards

| Old | New (`TEAMS_APP`) | Inbound activity |
|---|---|---|
| (manual `name === 'adaptiveCard/action'`) | `TEAMS_APP.on('card.action.<verb>', …)` | `AdaptiveCardInvokeActivity` |
| Generic catch-all | `TEAMS_APP.on('card.action', …)` | `AdaptiveCardInvokeActivity` |

#### Meetings

| Old (`TeamsActivityHandler`) | New (`TEAMS_APP.on(...)`) | Inbound activity |
|---|---|---|
| `handleTeamsMeetingStart` | `'meetingStart'` | `MeetingStartEventActivity` |
| `handleTeamsMeetingEnd` | `'meetingEnd'` | `MeetingEndEventActivity` |
| `handleTeamsMeetingParticipantsJoin` | `'meetingParticipantJoin'` | `MeetingParticipantJoinEventActivity` |
| `handleTeamsMeetingParticipantsLeave` | `'meetingParticipantLeave'` | `MeetingParticipantLeaveEventActivity` |

#### Top-level on `TeamsActivityHandler`

| Old | New (`TEAMS_APP.on(...)`) | Notes |
|---|---|---|
| `onTeamsMessageEdit` | `'message.update'` | `MessageUpdateActivity` |
| `onTeamsMessageUndelete` | `'message.update'` | `MessageUpdateActivity` (check `channelData.eventType`) |
| `onTeamsMessageSoftDelete` | `'message.delete'` | `MessageDeleteActivity` |
| `handleTeamsReadReceipt` | `'readReceipt'` | `ReadReceiptEventActivity` |
| `handleTeamsConfigFetch` | `'config.open'` | `ConfigFetchInvokeActivity` |
| `handleTeamsConfigSubmit` | `'config.submit'` | `ConfigSubmitInvokeActivity` |
| `handleTeamsFileConsentAccept` | `'file.consent'` (filter on `ctx.activity.value.action === 'accept'`) | `FileConsentInvokeActivity` |
| `handleTeamsFileConsentDecline` | `'file.consent'` (filter `... === 'decline'`) | same |
| `handleTeamsO365ConnectorCardAction` | *not exposed as a dedicated route* — use `TEAMS_APP.on('activity', …)` with a predicate, or fall through to `AGENT_SDK_APP` | |
| `onTeamsMembersAdded` / `onTeamsMembersRemoved` | conversation-update sub-events on `TEAMS_APP.on(...)` (e.g. `'teamMemberAdded'`); `'install.add'` / `'install.remove'` for the bot itself | `ConversationUpdateActivity` |
| `onTeamsChannelCreated` / `Deleted` / `Renamed` / `Restored` | `'channelCreated'` / `'channelDeleted'` / `'channelRenamed'` / `'channelRestored'` | `ConversationUpdateActivity` |
| `onTeamsTeamArchived` / `Deleted` / `HardDeleted` / `Renamed` / `Restored` / `Unarchived` | `'teamArchived'` / `'teamDeleted'` / `'teamHardDeleted'` / `'teamRenamed'` / `'teamRestored'` / `'teamUnarchived'` | `ConversationUpdateActivity` |

#### Other Teams SDK routes worth knowing about

These don't have a direct old-extension equivalent but are useful:

| New (`TEAMS_APP.on(...)`) | Purpose |
|---|---|
| `TEAMS_APP.message(pattern, cb)` | Inbound text messages — `pattern` is a `string` or `RegExp` matched against `activity.text` |
| `'messageReaction'` | Like / unlike on a previous message |
| `'message.submit'` | Thumbs up/down on AI-generated messages |
| `'typing'` | The user is typing |
| `'message.fetch-task'` | Action-based task fetch |
| `'signin.verify-state'` / `'signin.token-exchange'` | OAuth sign-in flows |
| `'activity'` | Any activity matching a custom predicate (use `TEAMS_APP.use(cb)` for cross-cutting middleware) |

> Anything not covered by a typed route can be reached via `TEAMS_APP.on('activity', …)` with your own discriminator, or — if no teams.ts route matches — the activity falls through to your `AgentApplication` handlers.

#### Task module response shape gotcha

The new Teams SDK uses **typed responses** that wrap differently than the old extension's plain-object responses:

```ts
// Old
return { task: { type: 'continue', value: { title: '…', card: { /* … */ } } } };

// New
import {
  TaskModuleInvokeResponse,
  TaskModuleContinueResponse,
  cardAttachment,
  AdaptiveCardAttachment,
} from '@microsoft/teams.api';

return new TaskModuleInvokeResponse({
  task: new TaskModuleContinueResponse({
    value: {
      title: 'Form',
      card: cardAttachment(new AdaptiveCardAttachment({ content: myCard() })),
    },
  }),
});
```

The `task` field on `TaskModuleInvokeResponse` (not `value`) is the common pitfall — the old extension accepted a `value` key and Teams would show "unable to reach app" because the response body was empty.

### Turn context

A single Teams turn now has **two contexts** in scope:

* **`ctx: ActivityContext<T>`** — the teams.ts context, passed to every `TEAMS_APP.on(...)` / `TEAMS_APP.message(...)` handler. Carries the typed activity, an `ApiClient` scoped to the inbound `serviceUrl`, and helpers like `ctx.send`, `ctx.reply`, `ctx.stream`, `ctx.api`.
* **`TurnContext`** — the Agents SDK context. Owns auth state, `turnState`, send hooks (`onSendActivities`), and is what every `AGENT_SDK_APP.on*` handler receives directly.

For most code you only need one of these and you'll use whichever is passed to your handler. If you're inside a `TEAMS_APP.*` handler and need the Agents SDK side (e.g. to read auth state or register a send hook):

```ts
import { agentSdkTurnContext } from 'teams-sdk-middleware';

TEAMS_APP.message('ping', async (ctx) => {
  const agentSdkCtx = agentSdkTurnContext();    // the Agents SDK TurnContext
  const userToken = agentSdkCtx.turnState.get('AccessToken');
  await ctx.send('pong');
});
```

`agentSdkTurnContext()` is backed by an `AsyncLocalStorage` that the middleware populates for the duration of the teams.ts handler. Calling it outside a Teams turn throws.

Going the other direction — from an `AGENT_SDK_APP.on*` handler to teams.ts functionality — see [Proactive flows](#proactive-flows).

### Proactive flows

A proactive flow is any send that *isn't* a direct reply to the current inbound activity — for example, sending an out-of-band notification after the turn completes, or sending from a webhook handler.

#### From a teams.ts handler — use `TEAMS_APP.send`

```ts
TEAMS_APP.message('notify me', async (ctx) => {
  const conversationId = ctx.activity.conversation.id;
  await ctx.send('Will ping you in 5s.');

  setTimeout(() => {
    TEAMS_APP.send(conversationId, { type: 'message', text: 'Ding!' })
      .catch((err) => console.error(err));
  }, 5000);
});
```

`TEAMS_APP.send` uses the API client constructed by `useTeamsSdk` — already wired to the Agents SDK's token provider.

#### From an `AGENT_SDK_APP.on*` handler — build a per-turn `ApiClient`

`TEAMS_APP.api` is pinned to a single `serviceUrl` at App construction time. Inbound activities through the Agents SDK can arrive with a *different* `serviceUrl` (e.g. `canary.botapi.skype.com/amer/...` vs `smba.trafficmanager.net/teams/`). Calling `TEAMS_APP.send` from an `AGENT_SDK_APP.on*` handler would route to the wrong region. The fix is a per-turn `ApiClient` against `context.activity.serviceUrl`, reusing the shared HTTP client (and therefore its token provider):

```ts
import { Client as ApiClient } from '@microsoft/teams.api';

AGENT_SDK_APP.onMessage('agents sdk react', async (context) => {
  const api = new ApiClient(context.activity.serviceUrl!, TEAMS_APP.api.http);
  await api.reactions.add(
    context.activity.conversation!.id,
    context.activity.id!,
    'like',
  );
});
```

This pattern also unlocks the full `api.reactions`, `api.meetings.*`, etc. surface from inside `AGENT_SDK_APP.on*` handlers.

> **Auto-typing note.** `AgentApplicationOptions.startTypingTimer` defaults to `true`. When you send via teams.ts's `ApiClient` directly (bypassing `context.sendActivity`), the Agents SDK's typing-stop hook never fires for that send. The actual message arrives correctly, but you may see a brief residual typing indicator. Set `startTypingTimer: false` if that matters.

### Reactive flows

A reactive flow is the normal "user sent something, bot responds" path. After installing the bridge, every reactive Teams turn does the following:

```
inbound activity (channel === 'msteams')
        │
        ▼
TeamsSdkMiddleware.onTurn
        │
        ├─ cast Activity → teams.ts Activity
        ├─ await TEAMS_APP.initialize()                    (idempotent)
        ├─ does TEAMS_APP.router select a handler?
        │     │
        │     ├─ matched → process via teams.ts
        │     │              ├─ TEAMS_APP.message('help', …) fires
        │     │              ├─ handler uses ctx.send / ctx.reply
        │     │              └─ invoke responses propagated to Agents SDK send pipeline
        │     │
        │     └─ unmatched → await next()
        │                    └─ AGENT_SDK_APP handlers run as usual
        │                        (e.g. AGENT_SDK_APP.onMessage('agents sdk proactive', …))
        ▼
turn complete
```

What this means for handler design:

* **Default to writing Teams-aware handlers on `TEAMS_APP`** — you get the typed activity, the rich builders, and the full Teams API.
* **Use `AGENT_SDK_APP` handlers for cross-channel or pure-Agents SDK behavior** — text patterns that should work on Teams, Webchat, and Slack alike; auth flows; OAuth callbacks.
* **Order doesn't matter.** Both handler sets are registered at startup; the middleware picks the right one per turn.

The bridge does not change anything about non-Teams turns. If `channelId !== 'msteams'`, `TeamsSdkMiddleware.onTurn` calls `next()` immediately and your `AgentApplication` runs untouched. Use this to support Teams + other channels from a single `AgentApplication` without conditional code paths.

---

## See also

* **Sample:** [`test_samples/teams-sdk-sample/`](../../test_samples/teams-sdk-sample/) — exercises reactive / proactive / invoke surfaces across both halves.
* **Teams SDK reference:** the standalone Teams SDK lives at [`microsoft/teams.ts`](https://github.com/microsoft/teams.ts); every `TEAMS_APP.on(...)` route name is documented there.
* **Bridge source:** [`libraries/teams-sdk-middleware/src/`](../../libraries/teams-sdk-middleware/src/) — the middleware, install helper, token provider, and `AsyncLocalStorage` wiring.
