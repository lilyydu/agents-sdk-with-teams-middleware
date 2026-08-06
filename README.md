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
(runnable samples) split:

| | Library (bridge) | Middleware sample | Vanilla sample |
|---|---|---|---|
| **python/** | `libraries/teams_sdk` | `test_samples/teams_sdk_sample` | `test_samples/teams_sdk_vanilla` |
| **typescript/** | `libraries/teams-sdk-middleware` | `test_samples/teams-sdk-sample` | `test_samples/teams-sdk-vanilla` |
| **dotnet/** | `libraries/TeamsSdkMiddleware` | `test_samples/TeamsMiddlewareSample` | — |

The **middleware** samples run an `AgentApplication` with the Teams SDK mounted through
the bridge, so both routers are live at once. The **vanilla** samples are the control
group: plain Teams SDK, no Agents SDK and no middleware. Run the same command against
both to see exactly what the bridge adds — and what it leaves untouched.

`tools/webchat/` is language-agnostic: a small Web Chat / Direct Line harness for
exercising whichever sample is currently bound to the bot endpoint on a **non-Teams**
channel. Together with the **Email** channel it covers the passthrough half of the
contract.

## Getting started

Build/run instructions per language, plus the end-to-end Teams setup (dev tunnel +
bot registration), are in **[AGENTS.md](./AGENTS.md)**. Each sample also has its own
`README.md`, and a `MIGRATION.md` covering the move from the previous handler-based
approach.

Quick pointers:
- **TypeScript:** `cd typescript && npm install && npm run build`
- **Python:** `cd python/test_samples/teams_sdk_sample && py -3.12 -m venv .venv && .venv\Scripts\python -m pip install -r requirements.txt` (Python 3.10+; see sample README for full steps)
- **.NET:** `cd dotnet && dotnet build TeamsMiddleware.slnx`
