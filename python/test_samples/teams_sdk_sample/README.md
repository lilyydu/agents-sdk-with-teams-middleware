# Teams SDK + Agents SDK Sample 

## Wiring

```python
from teams_sdk import use_teams_sdk

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

The sample's Teams SDK routes are `help`, `react`, `quote`, `targeted`, and
`task`. The Agents SDK handles `agents sdk react`, `agents sdk proactive`, and
the default echo fallback.


For every `msteams` turn the middleware checks whether `TEAMS_APP` has a
matching route; if so it hands the activity to
`TEAMS_APP.activity_processor.process_activity(...)` and propagates the
returned `InvokeResponse` back through the Agents SDK send pipeline so
invokes return their bodies correctly. If no teams.py route matches, the
turn falls through to `AGENT_SDK_APP`'s handlers.

## Reaching the Agents SDK `TurnContext` from a teams.py handler

```python
from teams_sdk import agent_sdk_turn_context

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
│  python/libraries/teams_sdk/                                     │
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

1. **Python 3.10+** — older versions won't work. On Windows, `py -3.12 --version` should report 3.12.x (or any 3.10+).
2. Create a venv inside this folder and install deps (this also editable-installs the local `teams_sdk` bridge):
   ```bash
   cd python/test_samples/teams_sdk_sample
   py -3.12 -m venv .venv                              # or `python3.12 -m venv .venv` on macOS/Linux
   .venv\Scripts\python -m pip install -U pip          # `.venv/bin/python` on macOS/Linux
   .venv\Scripts\python -m pip install -r requirements.txt
   ```
3. Drop the bot credentials in `.env` next to `app.py`:
   ```
   CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID=<guid>
   CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET=<secret>
   CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID=<guid>
   TOKENVALIDATION__ENABLED=false                      # optional; skip JWT validation for local dev
   PORT=3978
   ```
   For the sign-in demo, also add the OAuth handler entries shown under
   [User authentication](#user-authentication-two-graph-connections).
4. Start a dev tunnel pointing at `http://localhost:3978` and register/update a bot at `https://<tunnel>/api/messages` (e.g. `teams app create --name "..." --endpoint https://<tunnel>/api/messages --json`).
5. Run the bot:
   ```bash
   .venv\Scripts\python app.py
   ```
6. Install the bot in Teams and send `help` — replies are prefixed `[Teams SDK]` (Teams SDK route) or `[Agent SDK]` (fallthrough to `AgentApplication`).

## User authentication (two Graph connections)

`whoami` and `mail` both call Microsoft Graph, but through **separate OAuth connections** on the same AAD app:

| Command | Handler | ABS connection | Scopes | Graph call |
| --- | --- | --- | --- | --- |
| `whoami` | `graphuser` | `graphuser` | `User.Read` | `/me` |
| `mail` | `graphmail` | `graphmail` | `User.Read Mail.Read` | `/me/messages?$top=3` |

Each handler keeps its **own token cache**, so signing in for `whoami` does not satisfy
`mail` — the second command prompts its own sign-in, because `Mail.Read` was never consented
on the first token. That is the point of the demo.

Handlers are configured from the environment, not in code:

```
AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__graphuser__SETTINGS__AZUREBOTOAUTHCONNECTIONNAME=graphuser
AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__graphmail__SETTINGS__AZUREBOTOAUTHCONNECTIONNAME=graphmail
```

`app.py` forwards these by passing `**agents_sdk_config` into `AgentApplication`. `__TYPE` is
optional — `auth_type` defaults to `UserAuthorization`.

Provision each connection with a matching scope string:

```bash
az bot authsetting create --name <bot> --resource-group <rg> --setting-name graphmail \
  --client-id <aad-app> --client-secret <secret> --service Aadv2 \
  --provider-scope-string "User.Read Mail.Read" --parameters tenantId=<tenant>
```

### Auth only works on the Agents SDK side

`TeamsSDKMiddleware` calls `logic()` — the Agents SDK `on_turn` — only when no teams.py route
matches. The auth intercept lives *inside* that `on_turn`, so:

**a teams.py route can never be auth-protected.** It would not error; it would run
unauthenticated. `whoami` and `mail` are deliberately Agents SDK routes.

There is a second, sharper edge here. teams.py registers `signin/tokenExchange`,
`signin/verifyState` and `signin/failure` handlers *unconditionally* when the `App` is
constructed, so `router.select_handlers()` always matches them — including for a flow that
`AgentApplication` started. Those handlers verify against teams.py's own
`default_connection_name`, which defaults to `"graph"`. Left alone this means:

- the Agents SDK never learns its sign-in completed, so the command never replays, and
- if no ABS connection named `graph` exists, the lookup 404s and Teams reports
  **"unable to reach app"**.

The middleware therefore takes an optional `should_bypass_teams` predicate, evaluated only
for Teams turns, that forces a fall-through even when teams.py has a matching route. The
sample passes one that claims every `signin/*` invoke for `AgentApplication`:

```python
def _agent_sdk_owns_signin(context: TurnContext) -> bool:
    return context.activity.type == ActivityTypes.invoke and (
        context.activity.name or ""
    ).lower().startswith("signin/")


TEAMS_APP = use_teams_sdk(AGENT_SDK_APP, CONNECTION_MANAGER, _agent_sdk_owns_signin)
```

Drop the predicate if you want teams.py to own sign-in instead, and set
`default_connection_name` on the teams.py `App`.

### Channel support

| Channel | Sign-in |
| --- | --- |
| Teams | works — native OAuth card |
| Web Chat / Direct Line | works — sign-in link, then the original command replays |
| Email | **not supported** — declined up front |

Email is refused deliberately. Azure Bot Service flattens cards into a static image, so the
OAuth button is inert and the flow can never complete. Worse, once a flow is pending the auth
intercept swallows every later turn *before routing*, so the mailbox would go silent and even
`signout` could not recover it. Three higher-ranked routes (`RouteRank.FIRST`) match `whoami`,
`mail`, and `signout` on the email channel and decline before any flow starts.

`signout` is declined on email for a second reason: tokens are keyed by
`(channelId, userId, connectionName)`, and the email identity is an SMTP address — a separate
identity space from the Teams `29:…` id. That identity never holds a token, so signing out
there would report success for zero work and could never reach a token held on Teams.

The same "pending flow swallows everything" behaviour applies on any channel: abandon a
sign-in card and that user goes quiet until the flow expires. Restarting clears it, since the
sample uses `MemoryStorage`.

## Multichannel: Teams, Web Chat, and Email

`TeamsSDKMiddleware` routes to the teams.py `App` only when the activity is a Teams
activity; every other channel passes straight through to the Agents SDK app. Teams alone
can't show that half of the contract, so this sample is exercised on three channels.

| | Teams | Web Chat / Direct Line | Email |
| --- | --- | --- | --- |
| `channel` | `channelId=msteams (… fell through)` | `channelId=directline (… passed through)` | `channelId=email (… passed through)` |
| `help` | Adaptive Card via teams.py | plain-text help from the Agents SDK | plain-text help from the Agents SDK |
| `quote`, `task`, `react`, `targeted` | handled by teams.py | no teams.py route → echoed | no teams.py route → echoed |
| `agents sdk react` | uses the teams.py API client | politely declines — Teams-only API | politely declines — Teams-only API |
| `agents sdk proactive` | uses the teams.py API client | works — see below | works — see below |
| `whoami`, `mail` | sign-in via OAuth card | sign-in via OAuth card | declined — see [User authentication](#user-authentication-two-graph-connections) |
| `signout` | clears both handlers | clears both handlers | declined — nothing is ever signed in here |

### Web Chat / Direct Line

Direct Line is already enabled on the Azure Bot registration, so there is nothing extra to
provision. The repo ships a small harness at [`tools/webchat`](../../../tools/webchat) —
a browser UI plus a scriptable CLI:

```bash
python tools/webchat/serve.py                       # http://localhost:3000
python tools/webchat/dl_test.py help channel        # scripted, prints card contents too
```

### Email

Enable the **Email** channel on the bot registration and point it at a mailbox. Two things differ from every other channel:

- **`activity.text` is the message *body only*.** The subject arrives separately in
  `channelData.Subject`, so commands must go in the body — a subject-line command is
  invisible to the router.
- **The body carries a signature and/or quoted thread.** This is why `_command()` anchors
  matches to the *first line* rather than the whole body; without it, `help` followed by
  "Sent from my iPhone" would never match. The echo fallthrough truncates to the first
  line for the same reason.

Adaptive Cards do not fail on email — Azure Bot Service renders them **server-side into a
static image**. Display-only cards survive; a card with buttons becomes a painted, inert
control. That is why `help` has a separate plain-text handler on the Agents SDK side.

Mail sent to the bot mailbox from *outside* the tenant may be rejected on the reply leg, and
the resulting non-delivery report arrives back as an ordinary `message` activity that the
echo handler will answer. Send from an account in the same tenant to avoid this.
