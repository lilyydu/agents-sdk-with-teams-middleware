# Migration Guide (.NET)

> **Audience:** Existing developers using `TeamsActivityHandler` (the compat handler in `Microsoft.Agents.Extensions.Teams`) who want to move to the new Teams SDK middleware pattern shown in this sample — embedding the standalone Teams SDK (`Microsoft.Teams.Apps`) inside an Agents SDK `AgentApplication`.
>
> This is the .NET counterpart of the TypeScript migration guide. The shapes map almost 1:1; where .NET differs (DI registration, `IMiddleware[]`, `AsyncLocal`, experimental attributes) it is called out explicitly.

---

## Introduction

The Agents SDK ships a Teams "extension" in `Microsoft.Agents.Extensions.Teams`. The compat surface exposes Teams invokes (`task/fetch`, `composeExtension/query`, meeting lifecycle, …) as `protected virtual` methods you override on a `TeamsActivityHandler` subclass:

```csharp
using Microsoft.Agents.Extensions.Teams.Compat;

public class MyBot : TeamsActivityHandler
{
    protected override Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(
        ITurnContext<IInvokeActivity> turnContext,
        TaskModuleRequest request,
        CancellationToken cancellationToken)
    {
        // ...
    }
}
```

That surface covers Teams *invoke* shapes, but stops there. Things outside pure invoke handling — building rich activities (citations, AI labels, feedback, streaming, quotes), calling Teams APIs (reactions, conversation members, meeting participants), or sending proactively with full Teams metadata — were either missing, partial, or you had to drop down to a raw `ConnectorClient` / `TeamsInfo` and assemble the payloads yourself.

The **new** approach replaces the `TeamsActivityHandler` subclass with a thin **bridge middleware** (`TeamsSdkMiddleware`) that lets you embed a full `Microsoft.Teams.Apps.TeamsBotApplication` alongside your `AgentApplication`. You keep the Agents SDK as your hosting front door for non-Teams channels; for Teams turns the bridge hands the activity to the Teams SDK, so you get the full Teams SDK developer experience — typed activities, builders, an `ApiClient` with the full Teams surface, and rich routing — without giving up your existing `AgentApplication` handlers, auth, or storage.

```csharp
// Program.cs
builder.AddAgent<MyAgent>();
builder.Services.AddTeamsSdk<MyTeamsBot>();   // bot + Teams API/auth chain + routing middleware
```

```csharp
// MyTeamsBot.cs
public class MyTeamsBot : TeamsBotApplication
{
    public MyTeamsBot(ApiClient api, IHttpContextAccessor accessor, ILogger<MyTeamsBot> logger, TeamsBotApplicationOptions? options = null)
        : base(api, accessor, logger, options)
    {
        this.OnMessage("task", async (context, ct) => { /* ... */ });
    }
}
```

This document covers what changes, why, and how to migrate handler-by-handler.

---

## Motivation

| | Old `TeamsActivityHandler` | New `TeamsBotApplication` + bridge |
|---|---|---|
| **Handler surface** | Subset of Teams invokes, hand-curated `OnTeams*Async` overrides | Full Teams SDK — every activity, every invoke, every event, as fluent `On*` registrations |
| **Activity model** | Flat `Activity` + `activity.Value` (loosely typed) | Typed activity classes (`MessageActivity`, `TaskFetchInvokeActivity`, …) on `context.Activity` |
| **Outbound builders** | Manual objects + Adaptive Card JSON | `TeamsActivity.CreateBuilder()` with first-class citations, AI labels, feedback, streaming |
| **API client** | `ConnectorClient` + `TeamsInfo` for a few endpoints | Full `ApiClient` — `.Conversations.Reactions`, `.Conversations.Activities`, `.Meetings`, `.Users`, `.Teams` |
| **Streaming, citations, quotes** | Not supported / manual | First-class (`TeamsStreamingWriter`, `.AddCitation`, `context.Quote`) |
| **Release cadence** | Tied to `Microsoft.Agents.Extensions.Teams` | Teams SDK (`Microsoft.Teams.*`) ships independently — new Teams features arrive the moment teams.net releases them |
| **Hosting** | Agents SDK owns the HTTP endpoint, auth, storage | Same — Agents SDK still owns hosting; bridge is purely additive |

Net effect: Teams developers get the full Teams SDK feature set; Agents SDK developers keep their existing app, auth, and multi-channel surface. The bridge is one small `IMiddleware` (`TeamsSdkMiddleware.cs`) plus one DI helper (`TeamsSdkExtensions.cs`).

---

## Basic concepts

### What the middleware is and how it works

`TeamsSdkMiddleware` is a regular Agents SDK [`IMiddleware`](https://learn.microsoft.com/en-us/microsoft-365/agents-sdk/) registered on the `CloudAdapter` pipeline. Every turn passes through it; for **Teams turns** it acts as a router, for **other channels** it is a pure pass-through.

```
inbound HTTP /api/messages
        │
        ▼
CloudAdapter ── IMiddleware[] includes TeamsSdkMiddleware
        │
        ├─ ChannelId != Channels.Msteams ───────► AgentApplication handlers (unchanged)
        │
        └─ ChannelId == Channels.Msteams
                ├─ serialize Agents SDK IActivity → JSON → Teams SDK CoreActivity (same BF wire schema)
                ├─ stash ITurnContext in AsyncLocal (CurrentTurnContext)
                ├─ _teamsBot.HasMatchingRoute(activity)?
                │     ├─ yes →  invoke  → ProcessInvokeAsync → bridge InvokeResponse to Agents SDK send pipeline
                │     │         other  → _teamsBot.OnActivity(activity)
                │     │         then  return  (short-circuit; do NOT call next())
                │     └─ no  →  await next()  → AgentApplication handlers run as usual
```

Three points worth noting:

1. **The middleware never double-dispatches.** If a Teams SDK route matches (`HasMatchingRoute`), the `AgentApplication`'s own handlers do not also fire for that activity — the middleware `return`s without calling `next()`. If no route matches, the activity flows to `AgentApplication` exactly as if the bridge weren't installed.
2. **The route-match check uses a throwaway copy of the activity.** `HasMatchingRoute` calls `TeamsActivity.FromActivity`, which *mutates* the `CoreActivity` (its `Extract` calls remove entries from `Properties`). So the middleware deserializes a separate `CoreActivity` for the match check and a fresh one for the handler — see `TeamsSdkMiddleware.cs`. You don't have to think about this, but it's why the activity is deserialized twice.
3. **Invoke responses are propagated through the Agents SDK send pipeline.** Teams SDK invoke handlers return a typed `InvokeResponse`; the middleware wraps it via `Activity.CreateInvokeResponseActivity(...)` and sends it through `turnContext.SendActivityAsync` so the Agents SDK HTTP layer writes the synchronous body. Using `ProcessInvokeAsync` (rather than the non-invoke `OnActivity` path) avoids a double-write to `HttpContext.Response`.

### The Teams SDK and how it differs

The Teams SDK (`Microsoft.Teams.Apps`, `Microsoft.Teams.Core`, and its `Api`/`Cards` namespaces) is the Teams team's standalone .NET SDK for building Teams apps. It is **the same SDK you would use without the Agents SDK at all** — the bridge just lets you mount it on top of `AgentApplication`.

Three things to know:

#### Handlers

The Teams SDK registers handlers via fluent extension methods on the `TeamsBotApplication` instance — `OnMessage`, `OnQuery`, `OnTaskFetch`, `OnAdaptiveCardAction`, etc. There are no `protected virtual` methods to override:

```csharp
this.OnMessage(async (context, ct) => { /* any message */ });
this.OnMessage("help", async (context, ct) => { /* matches activity.Text */ });
this.OnTaskFetch(async (context, ct) => { /* TaskFetchInvokeActivity */ });
this.OnQuery(async (context, ct) => { /* message extension query */ });
```

Compared to `TeamsActivityHandler`:

* No `protected override OnTeams*Async` methods — every activity type and invoke shape has a registration method like `OnTaskFetch` or `OnQuery`.
* Handler signature is `(context, ct) => …` — a single typed `IContext<TActivity>` plus a `CancellationToken` — instead of `(ITurnContext<IInvokeActivity>, request, ct)`. The context carries the typed activity, an `ApiClient`, and helpers like `context.SendAsync`, `context.SendActivityAsync`, `context.Quote`.
* You don't unwrap `activity.Value` by hand — the typed activity on `context.Activity` already exposes the parsed fields (e.g. `context.Activity.Value.CommandId`).

See **Handler mapping** below for the one-to-one cheat sheet.

#### Activity model

Where the old extension gives you `IInvokeActivity` with a loosely typed `Value`, the Teams SDK models each shape as a typed activity. Outbound activities use **builders** that compose features cleanly:

```csharp
var activity = TeamsActivity.CreateBuilder()
    .WithText("See [1] for details.")
    .AddCitation(1, new CitationAppearance
    {
        Name = "Docs",
        Abstract = "…",
        Url = new Uri("https://…"),
        Icon = CitationIcon.Text,
    })
    .AddAIGenerated()
    .AddFeedback()
    .Build();

await context.SendActivityAsync(activity, ct);
```

> **Experimental APIs.** Some surfaces (e.g. quoted replies) are gated behind experimental attributes. The sample suppresses the relevant diagnostic at the top of `MyTeamsBot.cs`:
> ```csharp
> #pragma warning disable ExperimentalTeamsQuotedReplies // Quote is experimental
> ```

#### Underlying clients

The Teams SDK exposes an `ApiClient` — a single object with the **full Teams API surface**, grouped logically:

| Client | Old equivalent | What it does |
|---|---|---|
| `api.Conversations.Activities.CreateAsync(convId, …)` | `ConnectorClient.Conversations.SendToConversationAsync` | Send, update, delete, reply to activities |
| `api.Conversations.Members(...)` | `TeamsInfo.GetTeamMembersAsync` | Roster / membership |
| `api.Conversations.Reactions.AddAsync / DeleteAsync` | (manual payload) | Add / remove message reactions |
| `api.Meetings` | `TeamsInfo.GetMeetingInfoAsync` | Meetings, participants |
| `api.Teams` | `TeamsInfo.GetTeamDetailsAsync` | Team metadata |
| `api.Users` | (manual Graph) | User profiles |

Inside a Teams SDK handler you usually use `context.Api` (already scoped to the inbound `serviceUrl`). Inside an `AgentApplication` handler you scope one per-turn via `_teamsBot.Api.ForServiceUrl(...)` — see [Proactive flows](#proactive-flows) below.

---

## Initial setup

Replace the old `TeamsActivityHandler` wiring with the bridge registration. The three pieces live in `Program.cs`, `TeamsSdkExtensions.cs`, and `TeamsSdkMiddleware.cs`.

### Before

```csharp
using Microsoft.Agents.Extensions.Teams.Compat;

public class MyBot : TeamsActivityHandler
{
    protected override Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(
        ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken ct) { /* ... */ }

    protected override Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(
        ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest request, CancellationToken ct) { /* ... */ }
}

// Program.cs
builder.AddAgent<MyBot>();
```

### After

```csharp
// Program.cs
builder.AddAgent<MyAgent>();                                   // your AgentApplication (non-Teams + fall-through)
builder.Services.AddSingleton<IStorage, MemoryStorage>();
builder.Services.AddAgentAspNetAuthentication(builder.Configuration);

// Teams SDK in one call: bot + Teams API/auth chain (see TeamsSdkExtensions.cs /
// AgentSdkAuthHandler.cs) + the routing middleware on the CloudAdapter pipeline.
builder.Services.AddTeamsSdk<MyTeamsBot>();
```

```csharp
// MyTeamsBot.cs
public class MyTeamsBot : TeamsBotApplication
{
    public MyTeamsBot(ApiClient api, IHttpContextAccessor accessor, ILogger<MyTeamsBot> logger, TeamsBotApplicationOptions? options = null)
        : base(api, accessor, logger, options)
    {
        this.OnQuery(async (context, ct) => { /* ... */ });
        this.OnTaskFetch(async (context, ct) => { /* ... */ });
    }
}
```

`AddTeamsSdk<T>()` (in `TeamsSdkExtensions.cs`) does the wiring:

1. **Reuses Agent SDK auth.** It registers an `AgentSdkAuthHandler` (a `DelegatingHandler`) on a named `HttpClient`, so outbound Teams SDK calls acquire Bearer tokens via the Agent SDK's `IConnections` / `IAccessTokenProvider` — the same `clientId` / `tenantId` your Agents SDK app is already configured with. No separate `AzureAd` config section.
2. **Builds the Teams SDK clients** (`ConversationClient`, `UserTokenClient`, `ApiClient`) on top of that authenticated `HttpClient`.
3. **Registers your `TeamsBotApplication` subclass** as a singleton (and under the base `TeamsBotApplication` type) so the middleware can resolve it.
4. **Installs the routing middleware.** It registers `TeamsSdkMiddleware` as an Agents SDK `IMiddleware` and the `IMiddleware[]` the `CloudAdapter` consumes — so you don't wire the pipeline by hand.

Everything else — your `AgentApplicationOptions`, `CloudAdapter`, storage, error handlers, message handlers, auth handlers, ASP.NET host — stays exactly as it is. The Agents SDK is still your hosting layer.

---

## Guides

### Handler mapping

The Teams SDK uses fluent `On*` registration methods on the `TeamsBotApplication` instance. The following tables map old `TeamsActivityHandler` overrides to the corresponding Teams SDK method.

#### Message extensions (`composeExtension/*`)

| Old (`TeamsActivityHandler` override) | New (`this.On…`) |
|---|---|
| `OnTeamsMessagingExtensionQueryAsync` | `OnQuery` |
| `OnTeamsMessagingExtensionSelectItemAsync` | `OnSelectItem` |
| `OnTeamsMessagingExtensionSubmitActionAsync` | `OnSubmitAction` |
| `OnTeamsMessagingExtensionAgentMessagePreviewEditAsync` | `OnSubmitAction` (filter on `context.Activity.Value.BotMessagePreviewAction == "edit"`) |
| `OnTeamsMessagingExtensionAgentMessagePreviewSendAsync` | `OnSubmitAction` (filter `… == "send"`) |
| `OnTeamsMessagingExtensionFetchTaskAsync` | `OnFetchTask` |
| `OnTeamsAppBasedLinkQueryAsync` | `OnQueryLink` |
| `OnTeamsAnonymousAppBasedLinkQueryAsync` | `OnAnonQueryLink` |
| `OnTeamsMessagingExtensionConfigurationQuerySettingUrlAsync` | `OnQuerySettingUrl` |
| `OnTeamsMessagingExtensionConfigurationSettingAsync` | `OnSetting` |
| `OnTeamsMessagingExtensionCardButtonClickedAsync` | `OnCardButtonClicked` |

#### Task modules

| Old | New |
|---|---|
| `OnTeamsTaskModuleFetchAsync` | `OnTaskFetch` |
| `OnTeamsTaskModuleSubmitAsync` | `OnTaskSubmit` |

See the [Task module response shape](#task-module-response-shape-gotcha) note below — the response builders changed.

#### Adaptive cards

| Old | New |
|---|---|
| `OnTeamsCardActionInvokeAsync` (`adaptiveCard/action`) | `OnAdaptiveCardAction` |

#### Meetings

| Old | New |
|---|---|
| `OnTeamsMeetingStartAsync` | `OnMeetingStart` |
| `OnTeamsMeetingEndAsync` | `OnMeetingEnd` |
| `OnTeamsMeetingParticipantsJoinAsync` | `OnMeetingParticipantJoin` (also `OnMeetingJoin`) |
| `OnTeamsMeetingParticipantsLeaveAsync` | `OnMeetingParticipantLeave` (also `OnMeetingLeave`) |

#### Conversation update / teams / channels

| Old | New |
|---|---|
| `OnTeamsMembersAddedAsync` / `OnTeamsMembersRemovedAsync` | `OnMembersAdded` / `OnMembersRemoved`; `OnTeamMemberAdded` / `OnTeamMemberRemoved` for team-scoped; `OnInstall` / `OnUnInstall` for the bot itself |
| `OnTeamsChannelCreatedAsync` / `Deleted` / `Renamed` / `Restored` | `OnChannelCreated` / `OnChannelDeleted` / `OnChannelRenamed` / `OnChannelRestored` (also `OnChannelShared` / `OnChannelUnshared` / `OnChannelMemberAdded` / `OnChannelMemberRemoved`) |
| `OnTeamsTeamArchivedAsync` / `Deleted` / `HardDeleted` / `Renamed` / `Restored` / `Unarchived` | `OnTeamArchived` / `OnTeamDeleted` / `OnTeamHardDeleted` / `OnTeamRenamed` / `OnTeamRestored` / `OnTeamUnarchived` |
| (generic conversation update) | `OnConversationUpdate` |

#### Other top-level handlers

| Old | New | Notes |
|---|---|---|
| `OnTeamsMessageEditAsync` | `OnMessageUpdate` | check `channelData.eventType` to distinguish edit |
| `OnTeamsMessageUndeleteAsync` | `OnMessageUpdate` | check `channelData.eventType == "undeleteMessage"` |
| `OnTeamsMessageSoftDeleteAsync` | `OnMessageDelete` | |
| `OnTeamsReadReceiptAsync` | `OnReadReceipt` | |
| `OnTeamsFileConsentAcceptAsync` / `DeclineAsync` | `OnFileConsent` (filter on `context.Activity.Value.Action`) | |
| `OnTeamsConfigFetchAsync` / `OnTeamsConfigSubmitAsync` | `OnInvoke` with a predicate on the invoke name | no dedicated config route |
| `OnTeamsTabFetchAsync` / `OnTeamsTabSubmitAsync` | `OnInvoke` with a predicate | no dedicated tab route |
| `OnTeamsO365ConnectorCardActionAsync` | `OnInvoke` / `OnEvent` with a predicate, or fall through to `AgentApplication` | not exposed as a dedicated route |
| `OnTeamsSigninVerifyStateAsync` | `OAuthFlow` (`OnSignInComplete` / `OnSignInFailure`) | OAuth has its own flow object — different shape than the old single override |

#### Other Teams SDK routes worth knowing about

These have no direct old-extension equivalent but are useful:

| New (`this.On…`) | Purpose |
|---|---|
| `OnMessage(pattern, cb)` / `OnMessage(Regex, cb)` | Inbound text messages matched against `activity.Text` |
| `OnMessage(cb)` | Any inbound message |
| `OnMessageReaction` / `OnMessageReactionAdded` / `OnMessageReactionRemoved` | Like / unlike on a previous message |
| `OnMessageSubmitFeedback` | Thumbs up/down on AI-generated messages |
| `OnMessageSubmitAction` | Action submit on a message |
| `OnSuggestedActionSubmit` | Suggested action button tapped |
| `OnMessageFetchTask` | Action-based task fetch on a message |
| `OnEvent` | Any event activity |
| `OnInvoke` | Any invoke matching a custom predicate — the catch-all for shapes without a dedicated route |

> Anything not covered by a dedicated method can be reached via `OnInvoke` / `OnEvent` with your own discriminator, or — if no Teams SDK route matches — the activity falls through to your `AgentApplication` handlers.

#### Task module response shape gotcha

The Teams SDK uses **builder-produced typed responses** instead of the old plain-object responses:

```csharp
// Old
return new TaskModuleResponse
{
    Task = new TaskModuleContinueResponse
    {
        Value = new TaskModuleTaskInfo { Title = "Form", Card = attachment }
    }
};

// New
return TaskModuleResponse.CreateBuilder()
    .WithType(TaskModuleResponseType.Continue)
    .WithTitle("Form")
    .WithCard(attachment)
    .WithHeight(TaskModuleSize.Medium)
    .WithWidth(TaskModuleSize.Medium)
    .Build();
```

Message-extension actions wrap a task module the same way, via `MessageExtensionActionResponse.CreateBuilder().WithTask(...)`. Invoke handlers that have nothing to return acknowledge with `InvokeResponse.Ok()`. See `MyTeamsBot.cs` (`OnTaskFetch`, `OnTaskSubmit`, `OnFetchTask`, `OnSubmitAction`, `OnAdaptiveCardAction`) for working examples.

### Turn context

A single Teams turn now has **two contexts** in scope:

* **`context` (the Teams SDK `IContext<TActivity>`)** — passed to every `On*` handler. Carries the typed activity, an `ApiClient` scoped to the inbound `serviceUrl` (`context.Api`), and helpers like `context.SendAsync`, `context.SendActivityAsync`, `context.Quote`.
* **`ITurnContext` (the Agents SDK context)** — owns auth state, `ITurnState`, send hooks, and is what every `AgentApplication` handler receives directly.

If you're inside a Teams SDK handler and need the Agents SDK side (e.g. to read turn metadata or send through the Agents SDK pipeline), the middleware stashes the `ITurnContext` in an `AsyncLocal` exposed as `TeamsSdkMiddleware.CurrentTurnContext`:

```csharp
this.OnMessage("turn context", async (context, ct) =>
{
    var agentCtx = TeamsSdkMiddleware.CurrentTurnContext;   // the Agents SDK ITurnContext
    if (agentCtx is null) { /* not a Teams turn */ return; }

    await context.SendAsync("[Teams SDK] via Teams SDK context", ct);
    await agentCtx.SendActivityAsync(MessageFactory.Text("[Agent SDK] via Agent SDK turn context"), ct);
});
```

> **Why `AsyncLocal` and not `IHttpContextAccessor`?** Non-invoke activities are processed on a background thread where `HttpContext` is unavailable. `AsyncLocal` flows within the same async context regardless of which thread the turn runs on, so `CurrentTurnContext` works for both invoke and non-invoke turns. (This is the .NET analogue of the TypeScript guide's `AsyncLocalStorage` / `agentSdkTurnContext()`.)

Going the other direction — from an `AgentApplication` handler into Teams SDK functionality — see [Proactive flows](#proactive-flows).

### Proactive flows

A proactive flow is any send that *isn't* a direct reply to the current inbound activity — for example, sending an out-of-band notification after the turn completes.

#### From a Teams SDK handler — use `TeamsBotApplication.SendAsync`

```csharp
this.OnMessage("proactive", async (context, ct) =>
{
    string conversationId = context.Activity.Conversation!.Id;
    await context.SendAsync("You'll get a proactive message in ~3s...", ct);

    _ = Task.Run(async () =>
    {
        await Task.Delay(3000);
        await SendAsync(conversationId, "This is a proactive message sent outside the turn!");
    });
});
```

`SendAsync` uses the `ApiClient` constructed by `AddTeamsSdk` — already wired to the Agent SDK's token provider.

#### From an `AgentApplication` handler — scope an `ApiClient` to the inbound service URL

Inbound activities through the Agents SDK can arrive with different `serviceUrl`s (e.g. `canary.botapi.skype.com/amer/...` vs `smba.trafficmanager.net/teams/`). Use `_teamsBot.Api.ForServiceUrl(...)` to pin a per-turn client to `turnContext.Activity.ServiceUrl`, reusing the shared (authenticated) HTTP client:

```csharp
private async Task HandleAgentsReactAsync(ITurnContext turnContext, CancellationToken ct)
{
    var response = await turnContext.SendActivityAsync(MessageFactory.Text("…"), ct);

    var api = _teamsBot.Api.ForServiceUrl(new Uri(turnContext.Activity.ServiceUrl));
    string conversationId = turnContext.Activity.Conversation.Id;

    await api.Conversations.Reactions.AddAsync(conversationId, response.Id, ReactionTypes.Like, cancellationToken: ct);
    await api.Conversations.Reactions.DeleteAsync(conversationId, response.Id, ReactionTypes.Like, cancellationToken: ct);
}
```

This pattern also unlocks the full `api.Conversations.Activities`, `api.Meetings.*`, etc. surface from inside `AgentApplication` handlers — see `MyAgent.cs` (`HandleAgentsReactAsync`, `HandleAgentsProactiveAsync`, `HandleAgentsCitationAsync`) for the three working examples.

### Reactive flows

A reactive flow is the normal "user sent something, bot responds" path. After installing the bridge, every reactive Teams turn does the following:

```
inbound activity (ChannelId == Channels.Msteams)
        │
        ▼
TeamsSdkMiddleware.OnTurnAsync
        │
        ├─ serialize IActivity → CoreActivity
        ├─ stash ITurnContext in AsyncLocal (CurrentTurnContext)
        ├─ _teamsBot.HasMatchingRoute(activity)?
        │     │
        │     ├─ matched → process via Teams SDK
        │     │              ├─ invoke → ProcessInvokeAsync → bridge InvokeResponse to send pipeline
        │     │              └─ other  → _teamsBot.OnActivity(activity) (uses ConversationClient)
        │     │              return (short-circuit)
        │     │
        │     └─ unmatched → await next()
        │                    └─ AgentApplication handlers run (e.g. OnMessageAsync, rank: RouteRank.Last)
        ▼
turn complete
```

What this means for handler design:

* **Default to writing Teams-aware handlers on `MyTeamsBot`** — you get the typed activity, the rich builders, and the full Teams API.
* **Use `MyAgent` (`AgentApplication`) handlers for cross-channel or pure-Agents SDK behavior** — text patterns that should work on Teams, Webchat, and other channels alike; auth flows; OAuth callbacks; the default echo.
* **Order doesn't matter.** Both handler sets are registered at startup; the middleware picks the right one per turn based on `HasMatchingRoute`.

The bridge does not change anything about non-Teams turns. If `ChannelId != Channels.Msteams`, `OnTurnAsync` calls `next()` immediately and your `AgentApplication` runs untouched. Use this to support Teams + other channels from a single `AgentApplication` without conditional code paths.

---

## See also

* **This sample:** the files referenced throughout — `Program.cs` (wiring), `TeamsSdkExtensions.cs` (DI + auth bridge), `AgentSdkAuthHandler.cs` (token bridge), `TeamsSdkMiddleware.cs` (the router), `MyTeamsBot.cs` (Teams SDK handlers), `MyAgent.cs` (Agent SDK handlers using Teams SDK services). `COMMANDS.md` lists the demo commands.
* **Teams SDK reference:** the standalone Teams SDK lives at `microsoft/teams.net`; every `On*` registration method is defined under `src/Microsoft.Teams.Apps/Handlers/`.
* **TypeScript counterpart:** this guide mirrors the teams.ts `useTeamsSdk` migration guide; the concepts map 1:1 — `TeamsSdkMiddleware` is the shared name across both, `AddTeamsSdk<T>()` corresponds to `useTeamsSdk(...)`, and `TeamsSdkMiddleware.CurrentTurnContext` to `agentSdkTurnContext()`.
