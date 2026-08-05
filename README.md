# Teams SDK middleware for the Agents SDK

Prototypes of a **Teams SDK extension for the Microsoft Agents SDK**, in three
languages. Each embeds the standalone Teams SDK (`teams.py` / `teams.ts` /
`Microsoft.Teams.Apps`) inside an Agents SDK `AgentApplication` via middleware, so a
single bot registration, endpoint, and credential set serves both SDKs:

- **Teams turns with a matching Teams SDK route** → handled by the Teams SDK.
- **Everything else** (unmatched Teams turns, non-Teams channels) → falls through to
  the Agents SDK.

The three implementations are deliberately near-identical ports of the same design;
the core piece is the middleware.

## Layout

Each language tree has the same `libraries/` (the reusable bridge) + `test_samples/`
(a runnable sample) split:

| | Library (bridge) | Sample |
|---|---|---|
| **python/** | `libraries/teams_sdk` | `test_samples/teams_sdk_sample` |
| **typescript/** | `libraries/teams-sdk-middleware` | `test_samples/teams-sdk-sample` |
| **dotnet/** | `libraries/TeamsSdkMiddleware` | `test_samples/TeamsMiddlewareSample` |

## Getting started

Build/run instructions per language, plus the end-to-end Teams setup (dev tunnel +
bot registration), are in **[AGENTS.md](./AGENTS.md)**. Each sample also has its own
`README.md`, and a `MIGRATION.md` covering the move from the previous handler-based
approach.

Quick pointers:
- **TypeScript:** `cd typescript && npm install && npm run build`
- **Python:** `cd python/test_samples/teams_sdk_sample && py -3.12 -m venv .venv && .venv\Scripts\python -m pip install -r requirements.txt` (Python 3.10+; see sample README for full steps)
- **.NET:** `cd dotnet && dotnet build TeamsMiddleware.slnx`
- **Non-Teams harness:** `tools\webchat\` contains a shared Direct Line / Web Chat harness for exercising the Agents SDK fallthrough path.
