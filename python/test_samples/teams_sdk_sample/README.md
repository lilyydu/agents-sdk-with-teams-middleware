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
4. Start a dev tunnel pointing at `http://localhost:3978` and register/update a bot at `https://<tunnel>/api/messages` (e.g. `teams app create --name "..." --endpoint https://<tunnel>/api/messages --json`).
5. Run the bot:
   ```bash
   .venv\Scripts\python app.py
   ```
6. Install the bot in Teams and send `help` — replies are prefixed `[Teams SDK]` (Teams SDK route) or `[Agent SDK]` (fallthrough to `AgentApplication`).

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
