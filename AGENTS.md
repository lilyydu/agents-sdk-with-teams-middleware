# AGENTS.md

## What this repo is

Three parallel **prototypes of the Teams SDK extension for the Agents SDK** — one
per language. Each embeds the standalone Teams SDK (`teams.py` / `teams.ts` /
`Microsoft.Teams.Apps`) inside an Agents SDK `AgentApplication` so a single bot
registration, endpoint, and credential set serves both SDKs. The Agents SDK owns
hosting, auth, and any non-Teams channel; the Teams SDK owns Teams turns that
match one of its routes.

The three implementations are **deliberately near-identical** ports of the same
design. When changing one, mirror the change in the other two unless there's a
language-specific reason not to.

All three now follow the same `libraries/` (the reusable extension) + `test_samples/`
(a runnable sample that references it) split:

```
python/      libraries/teams_sdk                       +  test_samples/teams_sdk_sample
typescript/  libraries/teams-sdk-middleware            +  test_samples/teams-sdk-sample   (npm workspaces)
dotnet/      libraries/TeamsSdkMiddleware              +  test_samples/TeamsMiddlewareSample  (TeamsMiddleware.slnx)
```

## The core file: the middleware

The middleware is the heart of every prototype. Read it first:

- `python/libraries/teams_sdk/middleware.py` — `TeamsSDKMiddleware`
- `typescript/libraries/teams-sdk-middleware/src/middleware.ts` — `TeamsSdkMiddleware`
- `dotnet/libraries/TeamsSdkMiddleware/TeamsSdkMiddleware.cs` — `TeamsSdkMiddleware`

It implements the Agents SDK middleware interface (`on_turn` / `onTurn` /
`OnTurnAsync`) with the same turn lifecycle in all three:

1. **Non-Teams channel** (`channel_id != "msteams"`) → pass through to the
   Agents SDK (`logic()` / `next()`).
2. **Teams turn, no matching Teams SDK route** → pass through too.
3. **Teams turn with a match** → initialize the Teams app (idempotent), expose
   the Agents SDK `TurnContext` to Teams handlers via an ambient store
   (ContextVar / AsyncLocalStorage / AsyncLocal), hand the activity to the Teams
   SDK's processor, then for `invoke` activities propagate the returned
   `InvokeResponse` back through the Agents SDK send pipeline so the HTTP layer
   writes the synchronous body. It does **not** call `next()` afterward —
   doing so would run the Agents SDK handlers a second time.

Activity bridging between the two SDK models is done via JSON round-trip
(Python/.NET, both Pydantic/Activity-Protocol) or a structural cast (TypeScript,
shared camelCase wire shape).

## Supporting pieces (each language has equivalents)

- **Install helper** — one call that extracts `client_id`/`tenant_id` from the
  connection manager, wires the Teams SDK's outbound token callback to it,
  constructs the Teams app, and registers the middleware. `use_teams_sdk` (py),
  `useTeamsSdk` (ts), `AddTeamsSdk` / `TeamsSdkExtensions` (.NET).
- **Turn-context accessor** — `agent_sdk_turn_context()` (py),
  `agentSdkTurnContext` (ts), `TeamsSdkMiddleware.CurrentTurnContext` /
  `RequireTurnContext()` (.NET).
- **Credentials bridge** — adapts the Agents SDK connection manager / `IConnections`
  to the Teams SDK's token shape. `credentials.py`, `credentials.ts`,
  `AgentSdkAuthHandler.cs`.
- **Synthetic token** — projects the inbound Activity into the Teams SDK's
  `TokenProtocol` (the Agents SDK has already validated the JWT). `_token.py`,
  `token.ts`.

## Building & running the samples

> This is a freshly combined branch (`kavin/combine-all-sdks`); the three trees
> live under `python/`, `typescript/`, `dotnet/`, each with the `libraries/` +
> `test_samples/` split. Each sample consumes the Agents SDK and Teams SDK from its
> language's public package feed (npm / PyPI / nuget.org) — nothing is vendored.
> All three build, and all three have been run e2e against the same bot
> registration + dev tunnel (each binds `:3978` in turn).

### TypeScript (npm workspaces)
```bash
cd typescript
npm install
npm run build            # builds all workspaces (verified clean)
npm run start --workspace teams-sdk-sample   # needs typescript/test_samples/teams-sdk-sample/.env
```
Deps are public npm: `@microsoft/agents-*`, `@microsoft/teams.*`.

### Python
```bash
cd python/test_samples/teams_sdk_sample
python -m venv .venv
.venv/Scripts/python.exe -m pip install \
  aiohttp python-dotenv \
  microsoft-agents-activity microsoft-agents-hosting-core \
  microsoft-agents-hosting-aiohttp microsoft-agents-authentication-msal \
  microsoft-teams-apps microsoft-teams-api          # all on public PyPI
# Put the local teams library on the path WITHOUT pip-installing it: its setup.py
# pins microsoft-agents-hosting-core==<pkgver> (0.0.0 fallback) which won't resolve.
# microsoft_agents is a PEP 420 namespace package, so a .pth pointing at the lib
# root lets `microsoft_agents.hosting.teams` load from source while the rest of
# microsoft_agents.* resolves from site-packages:
echo "<abs>\python\libraries\microsoft-agents-hosting-teams" > .venv/Lib/site-packages/teams_local_lib.pth
python app.py            # reads .env; see env-var shape below
```
The Agents SDK reads nested config from env vars split on `__`. The sample needs a
`CONNECTIONS.SERVICE_CONNECTION` config — minimal `.env`:
```
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__AUTHTYPE=ClientSecret
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID=<guid>
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET=<secret>
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID=<guid>
```
(SETTINGS keys are `CLIENTID`/`CLIENTSECRET`/`TENANTID` — no underscores; `AgentAuthConfiguration`
maps them.) Serves on `http://localhost:3978/api/messages`; unauthenticated POSTs get
`401 {"error":"Authorization header not found"}` from `jwt_authorization_middleware`.

### .NET
Two projects, tied together by `dotnet/TeamsMiddleware.slnx`:
- `dotnet/libraries/TeamsSdkMiddleware/` — class library (namespace `TeamsSdkMiddleware`),
  the reusable bridge: `TeamsSdkMiddleware`, `AgentSdkAuthHandler`, `TeamsSdkExtensions`.
- `dotnet/test_samples/TeamsMiddlewareSample/` — web app (`Program.cs`, `MyAgent`,
  `MyTeamsBot`, `AspNetExtensions`), references the library via `<ProjectReference>`.

```bash
cd dotnet
dotnet build TeamsMiddleware.slnx           # builds library + sample (verified: 0 warn / 0 err)
dotnet run --project test_samples/TeamsMiddlewareSample --urls http://localhost:3978
```
Both csprojs reference the **published** packages — Agents SDK `Microsoft.Agents.*`
(`1.6.109-beta`, prerelease only) from nuget.org, Teams SDK `Microsoft.Teams.Apps`
(preview) from the `TeamsSDKPreviews` feed. `dotnet/nuget.config` (clear + source
mapping: `Microsoft.Teams.*` → preview feed, everything else → nuget.org) is scoped
to the `dotnet/` tree so it covers both projects without touching the repo root.
No Central Package Management / `Directory.Build.props` — each csproj is self-contained.

The middleware is decoupled from the sample: it depends on the base
`Microsoft.Teams.Apps.TeamsBotApplication`, and `AddTeamsSdk<T>`
registers the bot under that base type so the middleware resolves it regardless of
the concrete subclass. The library needs `<FrameworkReference Include="Microsoft.AspNetCore.App" />`
(it uses `IHttpContextAccessor` / DI extensions).

Config lives in the sample's `appsettings.json` (`Connections`, `TokenValidation`)
with `{{ClientId}}`-style placeholders; override at runtime with env vars
(`TokenValidation__Enabled`, `Connections__ServiceConnection__Settings__ClientId`, …)
to avoid editing the tracked file. `GET /` returns a version banner; in
non-Development with `TokenValidation:Enabled=true`, unauthenticated POSTs get `401`.

## Running a sample end-to-end in Teams

The Teams CLI + dev tunnels drive a real round-trip. Reference docs (Teams SDK,
LLM-optimized — do not web-search beyond these):
- Index: https://microsoft.github.io/teams-sdk/llms_docs/llms.txt
- Per language: `.../llms_csharp.txt`, `.../llms_typescript.txt`, `.../llms_python.txt`
  (append `_full` for the complete docs)
- Bot infra guide: https://microsoft.github.io/teams-sdk/llms_docs/references/guide-create-bot-infra.md
- Others: `guide-create-bot-app.md`, `guide-integrate-existing-server.md`,
  `guide-setup-sso.md`, `troubleshooting.md`

Validated flow (used for the .NET sample; `teams`/`devtunnel`/`az` must be logged in
— check `teams status`):
```bash
# 1. Public HTTPS endpoint via a dev tunnel on the sample's port (3978)
devtunnel create <name> --allow-anonymous
devtunnel port create <name> -p 3978 --protocol auto
devtunnel host <name>                       # long-running; prints https://<id>-3978.<cluster>.devtunnels.ms

# 2. Register a Teams-managed bot pointed at the tunnel (returns creds + installLink)
teams app create --name "TeamsMiddlewareSample" \
  --endpoint "https://<id>-3978.<cluster>.devtunnels.ms/api/messages" --json
#   (C#: add --env appsettings.json to write a Teams section; or inject creds via env vars)

# 3. Run the app on :3978 with the real CLIENT_ID / CLIENT_SECRET / TENANT_ID
# 4. Verify the public path: GET https://<tunnel>/ → 200 banner; POST /api/messages no-auth → 401
# 5. Open the installLink, install in Teams, send "help" / "Hi" → "[Teams SDK] ..." reply
```
Step 5 (install + message in the Teams client) is the only manual step — it needs a
real Bot Framework JWT, which can't be minted locally. Endpoint changed (new tunnel)?
`teams app update <teamsAppId> --endpoint "<new>/api/messages"`.

## Conventions

- Channel constant is the literal `"msteams"` in all three.
- Keep the three implementations aligned in behavior and in doc-comment intent.
- Each sample echoes with an explicit `[Teams SDK]` / `[Agent SDK]` prefix so you
  can see which SDK handled a turn.
- See each tree's `README.md` and the samples' `MIGRATION.md` for the full wiring
  walkthrough and migration guide from the previous handler-based approach.
