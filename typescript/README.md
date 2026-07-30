# Teams SDK Middleware (TypeScript)

A bridge that embeds the standalone Teams SDK (`@microsoft/teams.apps`) inside
an Agents SDK `AgentApplication` (`@microsoft/agents-hosting`).

This is the TypeScript port of the Python proof-of-concept at
[`lilyydu/agents-sdk-with-teams-middleware`](https://github.com/lilyydu/agents-sdk-with-teams-middleware).

## Layout

```
libraries/
  teams-sdk-middleware/   # the bridge — what users import
test_samples/
  teams-sdk-sample/       # live sample wiring it up
    MIGRATION.md          # migration guide from TeamsActivityHandler
```

## Quick start

```bash
npm install
npm run build
```

To run the sample, copy `test_samples/teams-sdk-sample/.env.example` to `.env`,
fill in `clientId` / `clientSecret` / `tenantId`, then:

```bash
npm run start --workspace teams-sdk-sample
```

## What the bridge gives you

```ts
import { AgentApplication, MsalConnectionManager } from '@microsoft/agents-hosting';
import { useTeamsSdk } from 'teams-sdk-middleware';

const AGENT_SDK_APP = new AgentApplication({ storage: /* ... */ });
const TEAMS_APP = useTeamsSdk(AGENT_SDK_APP, new MsalConnectionManager());

// teams.ts routes — typed activity, ctx.api, ctx.send, ctx.reply, etc.
TEAMS_APP.message('help', async (ctx) => {
  await ctx.send({ type: 'message', text: 'hello from teams.ts' });
});

// AgentApplication still owns hosting, auth, and any non-Teams channel.
AGENT_SDK_APP.onActivity('message', async (context) => {
  await context.sendActivity(`Echo: ${context.activity.text}`);
});
```

The sample's Teams SDK routes are `help`, `react`, `quote`, `targeted`, and
`task`. The Agents SDK handles `agents sdk react`, `agents sdk proactive`, and
the default echo fallback.

See [`test_samples/teams-sdk-sample/MIGRATION.md`](./test_samples/teams-sdk-sample/MIGRATION.md)
for the full migration guide from `TeamsActivityHandler` to this bridge.
