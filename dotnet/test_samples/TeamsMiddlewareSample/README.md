# Teams SDK + Agents SDK Sample (.NET)

## Wiring

```csharp
using TeamsSdk;

var builder = WebApplication.CreateBuilder(args);

builder.AddAgent<MyAgent>();
builder.Services.AddSingleton<IStorage, MemoryStorage>();
builder.Services.AddAgentAspNetAuthentication(builder.Configuration);

// One call: registers MyTeamsBot + its Teams SDK service chain (ApiClient /
// ConversationClient / UserTokenClient) on a named HttpClient authed via the Agent
// SDK's IConnections (AgentSdkAuthHandler), and installs the routing middleware on
// the CloudAdapter pipeline.
builder.Services.AddTeamsSdk<MyTeamsBot>();

var app = builder.Build();
app.MapAgentApplicationEndpoints(requireAuth: !app.Environment.IsDevelopment());
app.Run();
```

```csharp
// Teams SDK handlers live on MyTeamsBot (a TeamsBotApplication subclass):
this.OnMessage("help", async (context, ct) => await context.SendAsync(/* card */, ct));

// Agent SDK handlers live on MyAgent (an AgentApplication subclass):
OnActivity(ActivityTypes.Message, OnMessageAsync, rank: RouteRank.Last);   // echo fallthrough
```

`AddTeamsSdk<MyTeamsBot>()` is the only call you need. It registers the Teams SDK bot
and wires its outbound HTTP through `AgentSdkAuthHandler`, which acquires Bearer tokens
from the Agent SDK's `IConnections` — so both SDKs share one bot registration and
credential set (`Connections` in `appsettings.json`). It also registers the bot under
its base `TeamsBotApplication` type and installs `TeamsSdkMiddleware` (plus the
`IMiddleware[]` the `CloudAdapter` consumes) so Teams turns are routed without you
wiring up the pipeline by hand.

For every `msteams` turn the middleware checks whether `MyTeamsBot` has a matching
route; if so it hands the activity to the Teams SDK and, for `invoke` activities,
propagates the returned `InvokeResponse` back through the Agent SDK send pipeline so
invokes return their bodies correctly. If no Teams SDK route matches, the turn falls
through to `MyAgent`'s handlers. Any non-`msteams` channel goes straight to `MyAgent`.

## Reaching the Agent SDK `ITurnContext` from a Teams SDK handler

```csharp
this.OnMessage("turn context", async (context, ct) =>
{
    var agentCtx = TeamsSdkMiddleware.CurrentTurnContext;
    await context.SendAsync("[Teams SDK] ...", ct);                       // Teams SDK pipeline
    await agentCtx.SendActivityAsync(
        MessageFactory.Text("[Agent SDK] ..."), ct);                      // Agent SDK pipeline
});
```

`TeamsSdkMiddleware.CurrentTurnContext` returns the live Agent SDK `ITurnContext`
that the middleware set up for the current turn. It's held in an `AsyncLocal<T>`, so it
flows within the same async context even for non-invoke activities (processed on a
background thread where `HttpContext` is unavailable). Reads/writes through `agentCtx`
flow through the Agent SDK exactly as in a native `AgentApplication` turn. Use
`RequireTurnContext()` when its absence is an error; outside a Teams SDK-handled turn
`CurrentTurnContext` is `null`.

## Architecture

```
┌──────────────────────────────────────────────────────────────────┐
│  dotnet/libraries/TeamsSdkMiddleware/                            │
│                                                                  │
│  TeamsSdkMiddleware  (Agent SDK IMiddleware)               │
│   • non-Teams channel:    → await next() (Agent SDK)             │
│   • Teams channel:                                               │
│       serialize Activity → Teams SDK CoreActivity (JSON)         │
│       if no Teams SDK route matches → await next()               │
│       else:                                                      │
│         set AsyncLocal(agent SDK ITurnContext)                   │
│         invoke turns:    ProcessInvokeAsync, then emit the       │
│           InvokeResponse so the HTTP layer writes the body       │
│         non-invoke turns: OnActivity                             │
│         return (do NOT call next())                              │
│                                                                  │
│  AddTeamsSdk<T>(services)                           │
│    → registers T + its Teams SDK ApiClient chain on an           │
│      HttpClient whose outbound auth is AgentSdkAuthHandler       │
│      (bridges IConnections → Bearer tokens); registers the bot   │
│      under its base TeamsBotApplication so the middleware        │
│      resolves it                                                 │
│                                                                  │
│  TeamsSdkMiddleware.CurrentTurnContext                     │
│    → Agent SDK ITurnContext for the current turn                 │
└──────────────────────────────────────────────────────────────────┘
                              │ uses (no core changes)
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│  Microsoft.Agents.Hosting.AspNetCore  (published package)        │
│  • CloudAdapter + IMiddleware[] pipeline                         │
│  • MapAgentApplicationEndpoints (/api/messages)                  │
│  • AddAgentAspNetAuthentication (Bot Framework JWT validation)   │
└──────────────────────────────────────────────────────────────────┘
```

## Running

1. Create a bot registration (Teams CLI: `teams app create --name "TeamsMiddlewareSample" --endpoint https://your-tunnel.devtunnels.ms/api/messages --json`).
2. Put the credentials in `appsettings.json` (`Connections` + `TokenValidation`, with `TokenValidation.Enabled = true`), or supply them as env vars (`Connections__ServiceConnection__Settings__ClientId`, `TokenValidation__Enabled`, …) to keep them out of the tracked file.
3. Start a dev tunnel pointing at `http://localhost:3978`.
4. From `dotnet/`: `dotnet run --project test_samples/TeamsMiddlewareSample --urls http://localhost:3978`.
5. Install the bot in Teams using the install link from step 1 and send `help` (or `Hi` for the `[Agent SDK]` echo fallthrough).
