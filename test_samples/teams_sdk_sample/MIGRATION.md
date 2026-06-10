# Migration Guide

> **Audience:** Existing developers using `TeamsAgentExtension` (or the older `TeamsActivityHandler`) from `microsoft-agents-hosting-teams` who want to move to the new Teams SDK middleware (`use_teams_sdk`), which embeds the standalone Teams SDK (`microsoft-teams-*`) inside an Agents SDK `AgentApplication`.

---

## Introduction

The Agents SDK historically shipped a Teams "extension" in `microsoft-agents-hosting-teams`. That extension exposed Teams invokes (`task/fetch`, `composeExtension/query`, meeting lifecycle, …) as decorator methods on a `TeamsAgentExtension` wrapper around your `AgentApplication`:

```python
from microsoft_agents.hosting.teams import TeamsAgentExtension

teams = TeamsAgentExtension(agent_app)

@teams.task_module.on_fetch("openForm")
async def fetch(context, state, request):
    ...
```

That surface covers Teams *invoke* shapes, but stops there. Things outside of pure invoke handling — building rich activities (citations, AI labels, feedback, streaming, quotes), calling Teams APIs (reactions, conversation members, meeting participants), or sending proactively with full Teams metadata — were either missing, partial, or you had to reach into raw `ConnectorClient`/`TeamsInfo` and assemble JSON yourself.

The **new** approach replaces `TeamsAgentExtension` with a thin **bridge middleware** that lets you embed a full `microsoft_teams.apps.App` (the standalone Teams SDK) alongside your `AgentApplication`. You keep the Agents SDK as your hosting front door for non-Teams channels; for Teams turns, the bridge hands the activity to the Teams SDK so you get the full Teams SDK developer experience — typed activities, builders, an `ApiClient` with the full Teams surface, and rich routing — without giving up your existing `AgentApplication` handlers, auth, or storage.

```python
from microsoft_agents.hosting.teams import use_teams_sdk

TEAMS_APP = use_teams_sdk(AGENT_APP, CONNECTION_MANAGER)

@TEAMS_APP.on_message_pattern("task")
async def task(ctx):
    ...
```

This document covers what changes, why, and how to migrate handler-by-handler.

---

## Motivation

| | Old `TeamsAgentExtension` | New `use_teams_sdk` bridge |
|---|---|---|
| **Handler surface** | Subset of Teams invokes, hand-curated | Full Teams SDK — every activity, every invoke, every event |
| **Activity model** | Flat `Activity` + `activity.value: dict` | Typed Pydantic models (`MessageActivity`, `TaskFetchInvokeActivity`, …) via `ActivityTypeAdapter` |
| **Outbound builders** | Manual dicts + Adaptive Card JSON | `MessageActivityInput().add_card().add_citation().add_ai_generated().add_feedback()` chains |
| **API client** | `ConnectorClient` + `TeamsInfo` helper for a few endpoints | Full `ApiClient` — `.reactions`, `.conversations.activities`, `.meetings`, `.users`, `.teams` |
| **Streaming, citations, quotes** | Not supported / manual | First-class |
| **Release cadence** | Tied to `microsoft-agents-hosting-teams` | Teams SDK ships independently — new Teams features available the moment teams.py releases them |
| **Hosting** | Agents SDK owns the HTTP endpoint, auth, storage | Same — Agents SDK still owns hosting; bridge is purely additive |

Net effect: Teams developers get the full Teams SDK feature set; Agents SDK developers keep their existing app, auth, and multi-channel surface. The bridge is ~150 lines of glue.

---

## Basic concepts

### What is the middleware and how it works

`TeamsSDKMiddleware` is a regular Agents SDK [`Middleware`](https://learn.microsoft.com/en-us/microsoftteams/platform/) installed on `AGENT_APP.adapter`. Every turn passes through it; for **Teams turns** it acts as a router, for **other channels** it is a pure pass-through.

```
inbound HTTP /api/messages
        │
        ▼
CloudAdapter ── AGENT_APP.adapter.use(TeamsSDKMiddleware)
        │
        ├─ channel != "msteams"  ───────────► AgentApplication handlers (unchanged)
        │
        └─ channel == "msteams"
                ├─ translate Agents SDK Activity → teams.py typed Activity
                ├─ initialize() teams.py App (idempotent)
                ├─ expose TurnContext via ContextVar
                ├─ TEAMS_APP route matches?
                │     ├─ yes → run teams.py handler; propagate invoke response
                │     └─ no  → fall through to AgentApplication handlers
```

Three points worth noting:

1. **The middleware never double-dispatches.** If a teams.py route matches, `AgentApplication`'s own handlers do not also fire for that activity. If no teams.py route matches, the activity flows to `AgentApplication` exactly as if the bridge weren't installed.
2. **Initialization is per-turn, idempotent.** You don't need an explicit `await TEAMS_APP.initialize()` in your startup code — the middleware does it lazily so that `AgentApplication` handlers can safely call into `TEAMS_APP` even on the first turn.
3. **Invoke responses are propagated through the Agents SDK send pipeline.** Teams SDK invoke handlers return typed responses (e.g. `MessagingExtensionResponse`); the middleware wraps them in an `invoke_response` activity so the Agents SDK HTTP layer writes the synchronous body. You don't have to think about this.

### Introduction to the Teams SDK and how it's different

The Teams SDK (`microsoft-teams-*`: `microsoft-teams-api`, `microsoft-teams-apps`, `microsoft-teams-cards`, `microsoft-teams-common`) is the Teams team's standalone Python SDK for building Teams apps. It is **the same SDK you would use without the Agents SDK at all** — the bridge just lets you mount it on top of `AgentApplication`.

Three things to know:

#### Handlers

The Teams SDK uses decorator-based routing on the `App` object. Every Teams activity type and invoke shape has its own decorator:

```python
from microsoft_teams.apps import App, ActivityContext
from microsoft_teams.api import MessageActivity, TaskFetchInvokeActivity

@TEAMS_APP.on_message
async def any_message(ctx: ActivityContext[MessageActivity]): ...

@TEAMS_APP.on_message_pattern("help")
async def help_cmd(ctx: ActivityContext[MessageActivity]): ...

@TEAMS_APP.on_dialog_open
async def fetch(ctx: ActivityContext[TaskFetchInvokeActivity]): ...

@TEAMS_APP.on_message_ext_query
async def query(ctx): ...
```

Compared to `TeamsAgentExtension`:

* No sub-namespaces — `task_module.on_fetch` becomes `on_dialog_open` (the Teams SDK uses the modern "dialog" name), `message_extension.on_query` becomes `on_message_ext_query`, `meeting.on_start` becomes `on_meeting_start`.
* Handler signature is `(ctx: ActivityContext[T])` — a single context object — instead of `(context, state, value)`. The context carries the typed activity, the API client, helpers like `ctx.send`, `ctx.reply`, `ctx.stream`, and access to the inbound activity model.
* You don't unwrap `activity.value` yourself — the typed activity class on `ctx.activity` already has the parsed fields.

See **Handler mapping** below for a one-to-one cheat sheet.

#### Activity model

The Teams SDK models every activity shape as a discriminated Pydantic union, parsed by `ActivityTypeAdapter`. Where the Agents SDK gives you `Activity(type="invoke", name="task/fetch", value={...})`, the Teams SDK gives you a `TaskFetchInvokeActivity` with `activity.value.data`, `activity.value.context`, etc. as typed attributes.

```python
# Agents SDK
async def fetch(context, state, request):
    command_id = (context.activity.value or {}).get("data", {}).get("commandId")

# Teams SDK
async def fetch(ctx: ActivityContext[TaskFetchInvokeActivity]):
    command_id = ctx.activity.value.data.get("commandId")
```

Outbound activities use **builders** (`MessageActivityInput`) that compose features cleanly:

```python
from microsoft_teams.api import MessageActivityInput
from microsoft_teams.api.models.entity import CitationAppearance

message = (
    MessageActivityInput(text="See [1] for details.")
    .add_citation(1, CitationAppearance(name="Docs", abstract="...", url="..."))
    .add_ai_generated()
    .add_feedback()
)
await ctx.send(message)
```

#### Underlying clients

The Teams SDK exposes `microsoft_teams.api.ApiClient` — a single object with the **full Teams API surface**, grouped logically:

| Client | Old equivalent | What it does |
|---|---|---|
| `api.conversations.activities(conv_id)` | `ConnectorClient.conversations.send_to_conversation` | Send, update, delete, reply to activities |
| `api.conversations.members(conv_id)` | `TeamsInfo.get_team_members` | Roster / membership |
| `api.reactions` | (manual JSON) | Add / remove message reactions |
| `api.meetings` | `TeamsInfo.get_meeting_info` | Meetings, participants |
| `api.teams` | `TeamsInfo.get_team_details` | Team metadata |
| `api.users` | (manual Graph) | User profiles |
| `api.bots` | n/a | Bot self-info |

Inside a teams.py handler you usually use `ctx.api` (auto-scoped to the inbound `service_url`). Inside an `AgentApplication` handler you build one per-turn — see [Proactive flows](#proactive-flows) below.

---

## Initial setup

Replace the old extension wiring with a single call to `use_teams_sdk` *after* you build your `AgentApplication`:

### Before

```python
from microsoft_agents.hosting.core import AgentApplication
from microsoft_agents.hosting.teams import TeamsAgentExtension

AGENT_APP = AgentApplication[TurnState](options=...)
teams = TeamsAgentExtension(AGENT_APP)

@teams.message_extension.on_query("search")
async def search(context, state, query): ...

@teams.task_module.on_fetch("openForm")
async def fetch(context, state, request): ...
```

### After

```python
from microsoft_agents.hosting.core import AgentApplication
from microsoft_agents.hosting.teams import use_teams_sdk

AGENT_APP = AgentApplication[TurnState](options=...)
TEAMS_APP = use_teams_sdk(AGENT_APP, CONNECTION_MANAGER)

@TEAMS_APP.on_message_ext_query
async def search(ctx): ...

@TEAMS_APP.on_dialog_open
async def fetch(ctx): ...
```

`use_teams_sdk` does three things:

1. **Extracts credentials** from `CONNECTION_MANAGER.get_default_connection_configuration()` — the same `CLIENT_ID` / `TENANT_ID` your Agents SDK app is already configured with.
2. **Bridges the token provider** — outbound calls from the Teams SDK use the Agents SDK's `MsalConnectionManager`, so you don't double-configure auth.
3. **Installs `TeamsSDKMiddleware`** on `AGENT_APP.adapter`.

Everything else — your `ApplicationOptions`, `CloudAdapter`, storage, `@AGENT_APP.error`, `@AGENT_APP.message`, your auth handlers, your aiohttp app — stays exactly as it is. The Agents SDK is still your hosting layer.

---

## Guides

### Handler mapping

#### Message extensions (`composeExtension/*`)

| Old (`teams.message_extension`) | New (`TEAMS_APP`) | Inbound activity |
|---|---|---|
| `on_query("cmd")` | `on_message_ext_query` | `MessageExtensionQueryInvokeActivity` |
| `on_select_item` | `on_message_ext_select_item` | `MessageExtensionSelectItemInvokeActivity` |
| `on_submit_action("cmd")` | `on_message_ext_submit` | `MessageExtensionSubmitActionInvokeActivity` |
| `on_agent_message_preview_edit` | `on_message_ext_submit` (filter on `ctx.activity.value.bot_message_preview_action == "edit"`) | same |
| `on_agent_message_preview_send` | `on_message_ext_submit` (filter `... == "send"`) | same |
| `on_fetch_task` | `on_message_ext_open` | `MessageExtensionFetchTaskInvokeActivity` |
| `on_query_link` | `on_message_ext_query_link` | `MessageExtensionQueryLinkInvokeActivity` |
| `on_anonymous_query_link` | `on_message_ext_anon_query_link` | `MessageExtensionAnonQueryLinkInvokeActivity` |
| `on_query_url_setting` | `on_message_ext_query_settings_url` | `MessageExtensionQuerySettingUrlInvokeActivity` |
| `on_configure_settings` | `on_message_ext_setting` | `MessageExtensionSettingInvokeActivity` |
| `on_card_button_clicked` | `on_message_ext_card_button_clicked` | `MessageExtensionCardButtonClickedInvokeActivity` |

#### Task modules → "dialogs"

The Teams SDK renames *task modules* to **dialogs** (matches the Teams JS SDK 2.x rename).

| Old (`teams.task_module`) | New (`TEAMS_APP`) | Inbound activity |
|---|---|---|
| `on_fetch("cmd")` | `on_dialog_open("cmd")` | `TaskFetchInvokeActivity` |
| `on_submit("cmd")` | `on_dialog_submit("cmd")` | `TaskSubmitInvokeActivity` |

The response shapes also changed — see the [Task module response shape](#task-module-response-shape-gotcha) note below.

#### Adaptive cards

| Old | New (`TEAMS_APP`) | Inbound activity |
|---|---|---|
| (manual `name == "adaptiveCard/action"`) | `on_card_action_execute("verb")` | `AdaptiveCardInvokeActivity` |
| Generic catch-all | `on_card_action` | `AdaptiveCardInvokeActivity` |

#### Meetings

| Old (`teams.meeting`) | New (`TEAMS_APP`) | Inbound activity |
|---|---|---|
| `on_start` | `on_meeting_start` | `MeetingStartEventActivity` |
| `on_end` | `on_meeting_end` | `MeetingEndEventActivity` |
| `on_participants_join` | `on_meeting_participant_join` | `MeetingParticipantJoinEventActivity` |
| `on_participants_leave` | `on_meeting_participant_leave` | `MeetingParticipantLeaveEventActivity` |

#### Top-level on `TeamsAgentExtension`

| Old | New (`TEAMS_APP`) | Notes |
|---|---|---|
| `on_message_edit` | `on_edit_message` (or `on_message_update`) | `MessageUpdateActivity` |
| `on_message_undelete` | `on_undelete_message` (or `on_message_update`) | `MessageUpdateActivity` |
| `on_message_soft_delete` | `on_soft_delete_message` | `MessageDeleteActivity` |
| (hard delete) | `on_message_delete` | `MessageDeleteActivity` |
| `on_read_receipt` | `on_read_receipt` | `ReadReceiptEventActivity` |
| `on_config_fetch` | `on_config_open` | `ConfigFetchInvokeActivity` |
| `on_config_submit` | `on_config_submit` | `ConfigSubmitInvokeActivity` |
| `on_file_consent_accept` | `on_file_consent` (filter on `ctx.activity.value.action == "accept"`) | `FileConsentInvokeActivity` |
| `on_file_consent_decline` | `on_file_consent` (filter `... == "decline"`) | same |
| `on_o365_connector_card_action` | *not exposed in Teams SDK* — use `@TEAMS_APP.on_invoke` with predicate, or fall through to `AGENT_APP` | |
| `on_members_added` / `on_members_removed` | `on_conversation_update` (check `ctx.activity.members_added` / `members_removed`); `on_install_add` / `on_install_remove` for the bot itself | `ConversationUpdateActivity` |
| `on_channel_created` / `on_channel_deleted` / `on_channel_renamed` / `on_channel_restored` | same names | `ConversationUpdateActivity` |
| `on_team_archived` / `on_team_deleted` / `on_team_hard_deleted` / `on_team_renamed` / `on_team_restored` / `on_team_unarchived` | same names | `ConversationUpdateActivity` |

#### Other Teams SDK decorators worth knowing about

These don't have a direct old-extension equivalent but are useful:

| New (`TEAMS_APP`) | Purpose |
|---|---|
| `on_message` / `on_message_pattern(re)` | Inbound text messages — `on_message_pattern` runs your regex against `activity.text` |
| `on_message_reaction` | Like / unlike on a previous message |
| `on_message_submit_feedback` | Thumbs up/down on AI-generated messages |
| `on_typing` | The user is typing |
| `on_tab_open` / `on_tab_submit` | Tab fetch/submit |
| `on_signin_verify_state` / `on_signin_token_exchange` / `on_signin_failure` | OAuth sign-in flows |
| `on_invoke(name=...)` | Any invoke not covered by a typed decorator |
| `on_activity(predicate=...)` | Any activity matching a custom predicate |

> Anything not covered by a typed decorator can be reached via `@TEAMS_APP.on_invoke` / `@TEAMS_APP.on_activity`, or — if no teams.py route matches — the activity falls through to your `AgentApplication` handlers.

#### Task module response shape gotcha

The new Teams SDK uses **typed responses** that wrap differently than the old extension's dict-based responses:

```python
# Old
return {"task": {"type": "continue", "value": {"title": "...", "card": {...}}}}

# New
from microsoft_teams.api import (
    TaskModuleInvokeResponse,
    TaskModuleContinueResponse,
)
from microsoft_teams.api.models.task_module import CardTaskModuleTaskInfo
from microsoft_teams.api.models.attachment import (
    AdaptiveCardAttachment, card_attachment,
)

return TaskModuleInvokeResponse(
    task=TaskModuleContinueResponse(
        value=CardTaskModuleTaskInfo(
            title="Form",
            card=card_attachment(AdaptiveCardAttachment(content=my_card())),
        )
    )
)
```

The `task=` field on `TaskModuleInvokeResponse` (not `value=`) is the common pitfall — `CustomBaseModel` silently drops unknown kwargs and Teams shows "unable to reach app" because the response body is empty.

### Turn context

A single Teams turn now has **two contexts** in scope:

* **`ctx: ActivityContext[T]`** — the teams.py context, passed to every `@TEAMS_APP.*` handler. Carries the typed activity, an `ApiClient` scoped to the inbound `service_url`, and helpers like `ctx.send`, `ctx.reply`, `ctx.stream`, `ctx.api`.
* **`TurnContext`** — the Agents SDK context. Owns auth state, `turn_state`, send hooks (`on_send_activities`), and is what every `@AGENT_APP.*` handler receives directly.

For most code you only need one of these and you'll use whichever is passed to your handler. If you're inside a `@TEAMS_APP.*` handler and need the Agents SDK side (e.g. to read auth state or register a send hook):

```python
from microsoft_agents.hosting.teams import agent_turn_context

@TEAMS_APP.on_message_pattern("ping")
async def ping(ctx):
    agent_ctx = agent_turn_context()    # the Agents SDK TurnContext
    user_token = agent_ctx.turn_state.get("AccessToken")
    await ctx.send("pong")
```

`agent_turn_context()` is backed by a `ContextVar` that the middleware sets for the duration of the teams.py handler. Calling it outside a Teams turn raises `LookupError`.

Going the other direction — from an `@AGENT_APP.*` handler to teams.py functionality — see [Proactive flows](#proactive-flows).

### Proactive flows

A proactive flow is any send that *isn't* a direct reply to the current inbound activity — for example, sending an out-of-band notification after the turn completes, or sending from a webhook handler.

#### From a teams.py handler — use `TEAMS_APP.send`

```python
@TEAMS_APP.on_message_pattern("notify me")
async def schedule(ctx):
    conv_id = ctx.activity.conversation.id
    await ctx.send("Will ping you in 5s.")

    async def later():
        await asyncio.sleep(5)
        await TEAMS_APP.send(conv_id, "Ding!")
    asyncio.create_task(later())
```

`TEAMS_APP.send` uses the API client constructed by `use_teams_sdk` — already wired to the Agents SDK's token provider.

#### From an `@AGENT_APP.*` handler — build a per-turn `ApiClient`

`TEAMS_APP.api` is pinned to a single `service_url` at App construction time. Inbound activities through the Agents SDK can arrive with a *different* `service_url` (e.g. `canary.botapi.skype.com/amer/...` vs `smba.trafficmanager.net/teams/`). Calling `TEAMS_APP.send` from an `@AGENT_APP.*` handler would route to the wrong region. The fix is a per-turn `ApiClient` against `context.activity.service_url`, reusing the shared HTTP client (and therefore its token provider):

```python
from microsoft_teams.api.clients.api_client import ApiClient
from microsoft_teams.api import MessageActivityInput

@AGENT_APP.message("agents proactive")
async def proactive(context: TurnContext, _state: TurnState):
    conv_id = context.activity.conversation.id
    api = ApiClient(
        service_url=context.activity.service_url,
        options=TEAMS_APP.api.http,    # share the HTTP client / auth
    )
    await api.conversations.activities(conv_id).create(
        MessageActivityInput().add_text("Hello from the Agents SDK side!")
    )
```

This pattern also unlocks the full `api.reactions`, `api.meetings.*`, etc. surface from inside `@AGENT_APP.*` handlers.

> **Auto-typing note.** `ApplicationOptions.start_typing_timer` defaults to `True`. When you send via teams.py's `ApiClient` directly (bypassing `context.send_activity`), the Agents SDK's typing-stop hook never fires for that send. The actual message arrives correctly, but you may see a brief residual typing indicator. Set `start_typing_timer=False` if that matters.

### Reactive flows

A reactive flow is the normal "user sent something, bot responds" path. After installing the bridge, every reactive Teams turn does the following:

```
inbound activity (channel == "msteams")
        │
        ▼
TeamsSDKMiddleware.on_turn
        │
        ├─ Activity → ActivityTypeAdapter.validate_json → typed teams.py Activity
        ├─ TEAMS_APP.initialize()                          (idempotent)
        ├─ TEAMS_APP.router.select_handlers(activity)?
        │     │
        │     ├─ matched → process via teams.py
        │     │              ├─ @TEAMS_APP.on_message_pattern("help") fires
        │     │              ├─ handler uses ctx.send / ctx.reply
        │     │              └─ invoke responses propagated to Agents SDK send pipeline
        │     │
        │     └─ unmatched → await logic()
        │                    └─ AGENT_APP handlers run as usual
        │                        (e.g. @AGENT_APP.message("agents proactive"))
        ▼
turn complete
```

What this means for handler design:

* **Default to writing Teams-aware handlers on `TEAMS_APP`** — you get the typed activity, the rich builders, and the full Teams API.
* **Use `AGENT_APP` handlers for cross-channel or pure-Agents SDK behavior** — text patterns that should work on Teams, Webchat, and Slack alike; auth flows; OAuth callbacks.
* **Order doesn't matter.** Both handler sets are registered at startup; the middleware picks the right one per turn.

The bridge does not change anything about non-Teams turns. If `channel_id != "msteams"`, `TeamsSDKMiddleware.on_turn` calls `logic()` immediately and your `AgentApplication` runs untouched. Use this to support Teams + other channels from a single `AgentApplication` without conditional code paths.

---

## See also

* **Sample:** [`test_samples/teams_sdk_sample/`](./) — exercises every reactive / proactive / invoke surface across both halves.
* **Teams SDK reference:** the standalone Teams SDK lives at [`microsoft/teams.py`](https://github.com/microsoft/teams.py); every `@TEAMS_APP.*` decorator is documented there.
* **Bridge source:** [`libraries/microsoft-agents-hosting-teams/microsoft_agents/hosting/teams/teams_sdk/`](../../libraries/microsoft-agents-hosting-teams/microsoft_agents/hosting/teams/teams_sdk/) — the middleware, install helper, token provider, and ContextVar wiring.
