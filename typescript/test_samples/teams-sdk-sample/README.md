# Teams SDK + Agents SDK Sample (TypeScript)

## Wiring

```ts
import { useTeamsSdk } from 'teams-sdk-middleware';

const AGENT_SDK_APP = new AgentApplication<TurnState>({ storage, adapter });
const TEAMS_APP = useTeamsSdk(AGENT_SDK_APP, ADAPTER.connectionManager);

TEAMS_APP.message('help', async (ctx) => { /* ... */ });

AGENT_SDK_APP.onActivity('message', async (context) => { /* ... */ });
```

`useTeamsSdk` extracts `clientId`/`tenantId` from the connection manager, wires teams.ts's
outbound token callback to it, constructs the `@microsoft/teams.apps` `App`, and installs
`TeamsSdkMiddleware` on the Agents SDK adapter — returning the configured `App` ready for
handler registration. Pass extra `App` constructor options (e.g. `logger`, `plugins`) in the
third argument.

The sample's Teams SDK routes are `help`, `react`, `quote`, `targeted`, and `task`. The
Agents SDK handles `channel`, `whoami`, `mail`, `signout`, `agents sdk react`,
`agents sdk proactive`, and the default echo fallback.

For every `msteams` turn the middleware checks whether `TEAMS_APP` has a matching route; if
so it hands the activity to `TEAMS_APP.process(...)` and propagates the returned
`InvokeResponse` back through the Agents SDK send pipeline so invokes return their bodies
correctly. If no teams.ts route matches, the turn falls through to `AGENT_SDK_APP`'s
handlers.

## Reaching the Agents SDK `TurnContext` from a teams.ts handler

```ts
import { agentSdkTurnContext } from 'teams-sdk-middleware';

TEAMS_APP.message('turn context', async (ctx) => {
  const agentSdkCtx = agentSdkTurnContext();
  await ctx.send('[Teams SDK] ...');
  await agentSdkCtx.sendActivity('[Agent SDK] ...');
});
```

`agentSdkTurnContext()` returns the live Agents SDK `TurnContext` that `TeamsSdkMiddleware`
set up for the current turn (via `AsyncLocalStorage`). Outside of a Teams SDK-handled turn
the store is unset and the helper throws.

## Running

1. **Node 20+.**
2. Install and build from the `typescript/` root (npm workspaces):
   ```bash
   cd typescript
   npm install
   npm run build
   ```
3. Drop the bot credentials in `test_samples/teams-sdk-sample/.env`:
   ```
   clientId=<guid>
   clientSecret=<secret>
   tenantId=<guid>
   PORT=3978
   ```
   For the sign-in demo, also add the OAuth handler entries shown under
   [User authentication](#user-authentication-two-graph-connections).
4. Start a dev tunnel pointing at `http://localhost:3978` and register/update a bot at
   `https://<tunnel>/api/messages` (e.g.
   `teams app create --name "..." --endpoint https://<tunnel>/api/messages --json`).
5. Run the bot:
   ```bash
   npm run start --workspace teams-sdk-sample
   ```
6. Install the bot in Teams and send `help` — replies are prefixed `[Teams SDK]` (Teams SDK
   route) or `[Agent SDK]` (fallthrough to `AgentApplication`).

## User authentication (two Graph connections)

`whoami` and `mail` both call Microsoft Graph, but through **separate OAuth connections** on
the same AAD app:

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

`AgentApplication` reads them straight from `process.env` at construction (`npm start` loads
`.env` via `node --env-file`). `__TYPE` is optional — it defaults to `UserAuthorization`.
`index.ts` derives `AUTH_HANDLER_IDS` from the same variables so `signout` knows what to
clear.

Provision each connection with a matching scope string:

```bash
az bot authsetting create --name <bot> --resource-group <rg> --setting-name graphmail \
  --client-id <aad-app> --client-secret <secret> --service Aadv2 \
  --provider-scope-string "User.Read Mail.Read" --parameters tenantId=<tenant>
```

### Auth only works on the Agents SDK side

`TeamsSdkMiddleware` calls `next()` — the Agents SDK turn — only when no teams.ts route
matches. The auth intercept lives *inside* that turn, so:

**a teams.ts route can never be auth-protected.** It would not error; it would run
unauthenticated. `whoami` and `mail` are deliberately Agents SDK routes.

There is a second, sharper edge here. teams.ts registers `signin/tokenExchange`,
`signin/verifyState` and `signin/failure` handlers *unconditionally* when the `App` is
constructed, so the router always matches them — including for a flow that
`AgentApplication` started. Those handlers verify against teams.ts's own
`defaultConnectionName`, which defaults to `"graph"`. Left alone this means:

- the Agents SDK never learns its sign-in completed, so the command never replays, and
- if no ABS connection named `graph` exists, the lookup 404s and Teams reports
  **"unable to reach app"**.

The middleware therefore takes an optional `shouldBypassTeams` predicate, evaluated only for
Teams turns, that forces a fall-through even when teams.ts has a matching route. The sample
passes one that claims every `signin/*` invoke for `AgentApplication`:

```ts
function agentSdkOwnsSignIn(context: TurnContext): boolean {
  return (
    context.activity.type === ActivityTypes.Invoke &&
    (context.activity.name ?? '').toLowerCase().startsWith('signin/')
  );
}

const TEAMS_APP = useTeamsSdk(AGENT_SDK_APP, CONNECTION_MANAGER, {}, agentSdkOwnsSignIn);
```

Drop the predicate if you want teams.ts to own sign-in instead, and set
`defaultConnectionName` on the teams.ts `App`.

### Channel support

| Channel | Sign-in |
| --- | --- |
| Teams | works — native OAuth card |
| Web Chat / Direct Line | works — sign-in link, then the original command replays |
| Email | **not supported** — declined up front |

Email is refused deliberately. Azure Bot Service flattens cards into a static image, so the
OAuth button is inert and the flow can never complete. Worse, once a flow is pending the auth
intercept swallows every later turn *before routing*, so the mailbox would go silent and even
`signout` could not recover it. Three higher-ranked routes (`RouteRank.First`) match
`whoami`, `mail`, and `signout` on the email channel and decline before any flow starts.

`signout` is declined on email for a second reason: tokens are keyed by
`(channelId, userId, connectionName)`, and the email identity is an SMTP address — a separate
identity space from the Teams `29:…` id. That identity never holds a token, so signing out
there would report success for zero work and could never reach a token held on Teams.

The same "pending flow swallows everything" behaviour applies on any channel: abandon a
sign-in card and that user goes quiet until the flow expires. Restarting clears it, since the
sample uses `MemoryStorage`.

## Multichannel: Teams, Web Chat, and Email

`TeamsSdkMiddleware` routes to the teams.ts `App` only when the activity is a Teams
activity; every other channel passes straight through to the Agents SDK app. Teams alone
can't show that half of the contract, so this sample is exercised on three channels.

| | Teams | Web Chat / Direct Line | Email |
| --- | --- | --- | --- |
| `channel` | `channelId=msteams (… fell through)` | `channelId=directline (… passed through)` | `channelId=email (… passed through)` |
| `help` | Adaptive Card via teams.ts | plain-text help from the Agents SDK | plain-text help from the Agents SDK |
| `quote`, `task`, `react`, `targeted` | handled by teams.ts | no teams.ts route → echoed | no teams.ts route → echoed |
| `agents sdk react` | uses the teams.ts API client | politely declines — Teams-only API | politely declines — Teams-only API |
| `agents sdk proactive` | uses the teams.ts API client | works — see below | works — see below |
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

Enable the **Email** channel on the bot registration and point it at a mailbox. Two things
differ from every other channel:

- **`activity.text` is the message *body only*.** The subject arrives separately in
  `channelData.Subject`, so commands must go in the body — a subject-line command is
  invisible to the router.
- **The body carries a signature and/or quoted thread.** This is why `command()` anchors
  matches to the *first line* rather than the whole body; without it, `help` followed by
  "Sent from my iPhone" would never match. The echo fallthrough truncates to the first line
  for the same reason.

Adaptive Cards do not fail on email — Azure Bot Service renders them **server-side into a
static image**. Display-only cards survive; a card with buttons becomes a painted, inert
control. That is why `help` has a separate plain-text handler on the Agents SDK side.

Mail sent to the bot mailbox from *outside* the tenant may be rejected on the reply leg, and
the resulting non-delivery report arrives back as an ordinary `message` activity that the
echo handler will answer. Send from an account in the same tenant to avoid this.

## Migration

See [`MIGRATION.md`](./MIGRATION.md) for the migration guide from `TeamsActivityHandler` to
this bridge.
