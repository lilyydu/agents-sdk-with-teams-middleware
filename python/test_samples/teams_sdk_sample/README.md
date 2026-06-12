# Teams SDK + Agents SDK Sample 

## Wiring

```python
from microsoft_agents.hosting.teams import use_teams_sdk, agent_sdk_turn_context

AGENT_SDK_APP = AgentApplication(...)
TEAMS_APP = use_teams_sdk(AGENT_SDK_APP, CONNECTION_MANAGER)

@TEAMS_APP.on_message_pattern("help")
async def _help(ctx): ...

@AGENT_SDK_APP.activity("message")
async def _echo(context, state): ...
```

`use_teams_sdk` extracts `client_id`/`tenant_id` from the connection
manager, wires teams.py's outbound token callback to it, constructs the
`microsoft_teams.apps.App`, and installs `TeamsSDKMiddleware` on the
Agents SDK adapter — returning the configured `App` ready for handler
registration. Pass extra `App` constructor options (e.g., `logger=`,
`plugins=`) as keyword args.

For every `msteams` turn the middleware checks whether `TEAMS_APP` has a
matching route; if so it hands the activity to
`TEAMS_APP.activity_processor.process_activity(...)` and propagates the
returned `InvokeResponse` back through the Agents SDK send pipeline so
invokes return their bodies correctly. If no teams.py route matches, the
turn falls through to `AGENT_SDK_APP`'s handlers.

## Reaching the Agents SDK `TurnContext` from a teams.py handler

```python
from microsoft_agents.hosting.teams import agent_sdk_turn_context

@TEAMS_APP.on_message_pattern("turn context")
async def _turn_ctx(ctx):
    agent_sdk_ctx = agent_sdk_turn_context()
    await ctx.send("[Teams SDK] ...")
    await agent_sdk_ctx.send_activity("[Agent SDK] ...")
```

`agent_sdk_turn_context()` returns the live Agents SDK `TurnContext` that
`TeamsSDKMiddleware` set up for the current turn (via a `ContextVar`).
Mutations to `agent_sdk_ctx.turn_state` and outbound sends through
`agent_sdk_ctx.send_activity` flow through the Agents SDK exactly as in a
native `AgentApplication` turn. Outside of a Teams SDK-handled turn the
ContextVar is unset and the helper raises `LookupError`.

## Architecture

```
┌──────────────────────────────────────────────────────────────────┐
│  microsoft-agents-hosting-teams/teams_sdk/                       │
│                                                                  │
│  TeamsSDKMiddleware                                              │
│   • non-Teams channel:    → await next() (Agents SDK)            │
│   • Teams channel:                                               │
│       translate Activity → teams.py Activity                     │
│       if no teams.py route matches → await next()                │
│       else:                                                      │
│         await TEAMS_APP.initialize() (idempotent)                │
│         set ContextVar(agent_sdk TurnContext)                    │
│         await TEAMS_APP.activity_processor.process_activity      │
│         reset ContextVar                                         │
│         for invoke turns: emit invoke_response activity so       │
│           ChannelServiceAdapter writes the synchronous body      │
│                                                                  │
│  use_teams_sdk(agent_sdk_app, connection_manager, **App kwargs)      │
│    → extracts client/tenant from connection_manager,             │
│      wires teams.py token callback to it (tenant-aware via       │
│      ClaimsIdentity read from turn_state[AGENT_IDENTITY_KEY],    │
│      falls back to default connection for proactive sends),      │
│      constructs the teams.py App,                                │
│      registers TeamsSDKMiddleware on agent_sdk_app.adapter,          │
│      returns the App for handler registration                    │
│                                                                  │
│  agent_sdk_turn_context()                                        │
│    → Agents SDK TurnContext for the current turn                 │
└──────────────────────────────────────────────────────────────────┘
                              │ uses (no core changes)
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│  microsoft-agents-hosting-core  (UNCHANGED)                      │
│  • adapter.use(middleware)                                       │
│  • ChannelServiceAdapter writes invoke responses stashed in      │
│    turn_state[INVOKE_RESPONSE_KEY] by outbound invoke_response   │
└──────────────────────────────────────────────────────────────────┘
```

## Running

1. Create a bot registration (Teams CLI: `teams app create --name "..." --endpoint https://your-tunnel.devtunnels.ms/api/messages --json`).
2. Drop the credentials in `.env` next to `app.py` (`CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID`).
3. Start a dev tunnel pointing at `http://localhost:3978`.
4. `python app.py` (with the repo's venv activated).
5. Install the bot in Teams and send `help`.
