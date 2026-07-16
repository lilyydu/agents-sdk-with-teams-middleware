# Teams SDK + Agents SDK Sample (.NET)

## Wiring

```csharp
using TeamsExtension;

var builder = WebApplication.CreateBuilder(args);

builder.AddAgent<MyAgent>();
builder.Services.AddSingleton<IStorage, MemoryStorage>();
builder.Services.AddAgentAspNetAuthentication(builder.Configuration);

// Registers MyTeamsBot + Teams SDK clients (ApiClient / ConversationClient /
// UserTokenClient) authenticated through Agent SDK IConnections.
builder.Services.AddTeamsExtension<MyTeamsBot>();

var app = builder.Build();
app.MapAgentApplicationEndpoints(requireAuth: !app.Environment.IsDevelopment());
app.Run();
```

```csharp
// Teams SDK handlers live on MyTeamsBot (TeamsBotApplication):
this.OnMessage("help", async (context, ct) => await context.SendAsync(/* card */, ct));

// Agent SDK handlers live on MyAgent (AgentApplication):
OnActivity(ActivityTypes.Message, OnMessageAsync, rank: RouteRank.Last);

// Bridge Teams routes into the same AgentApplication turn lifecycle:
this.UseTeamsExtension(teamsBot);
```

`AddTeamsExtension<MyTeamsBot>()` configures Teams SDK services and outbound auth.  
`UseTeamsExtension(...)` adds an AgentApplication extension that routes matching
`msteams` turns to `MyTeamsBot` in `OnBeforeTurn`, then short-circuits Agent routes
for handled activities. Unmatched Teams routes and non-Teams channels fall through to
`MyAgent`.

## Reaching Agent SDK `ITurnContext` from a Teams SDK handler

```csharp
this.OnMessage("turn context", async (context, ct) =>
{
    var agentCtx = TeamsSdkAgentExtension.CurrentTurnContext;
    await context.SendAsync("[Teams SDK] ...", ct);
    await agentCtx.SendActivityAsync(MessageFactory.Text("[Agent SDK] ..."), ct);
});
```

`TeamsSdkAgentExtension.CurrentTurnContext` exposes the current Agent SDK turn context
for Teams SDK-routed turns and is stored with `AsyncLocal<T>`.

## Running

1. Create a bot registration (Teams CLI: `teams app create --name "TeamsMiddlewareSample" --endpoint https://your-tunnel.devtunnels.ms/api/messages --json`).
2. Put credentials in `appsettings.json` (`Connections` + `TokenValidation`, with `TokenValidation.Enabled = true`), or supply env vars (`Connections__ServiceConnection__Settings__ClientId`, `TokenValidation__Enabled`, ...).
3. Start a dev tunnel pointing at `http://localhost:3978`.
4. From `dotnet/`: `dotnet run --project test_samples/TeamsMiddlewareSample --urls http://localhost:3978`.
5. Install the bot in Teams using the install link from step 1 and send `help` (or `Hi` for `[Agent SDK]` echo fallthrough).
