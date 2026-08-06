# Vanilla Teams SDK sample

Plain **teams.ts** —  it exists to demonstrate two things:

1. **Raw JSON activities.** Sending and inspecting activities as plain objects rather
   than typed models, and seeing exactly what the connector answers.
2. **Non-Teams channels.** teams.ts never inspects `channelId`, so the same handlers run on
   Direct Line and email. `dump` makes the differences visible.

## Setup

```bash
cd typescript
npm install
cp test_samples/teams-sdk-vanilla/.env.example test_samples/teams-sdk-vanilla/.env
#   then fill in CLIENT_ID / CLIENT_SECRET / TENANT_ID
npm run build --workspace teams-sdk-vanilla
npm run start --workspace teams-sdk-vanilla
```

teams.ts reads `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID` straight from the environment —
there is no configuration object and no `loadAuthConfigFromEnv`, unlike the Agents SDK
sample next door.

The app listens on **3979** so it can run alongside `teams-sdk-sample` on 3978. Point a tunnel
at it and set that URL as the bot's messaging endpoint:

```bash
devtunnel create vanilla-sample --allow-anonymous
devtunnel port create vanilla-sample -p 3979 --protocol http   # http, not https — the app serves plain HTTP
devtunnel host vanilla-sample

az bot update -n <botAppId> -g <rg> --endpoint "https://<tunnel>-3979.<region>.devtunnels.ms/api/messages"
```

An unauthenticated `POST /api/messages` returning **401 is healthy** — it means token
validation is running.

## Commands

| Command | What it does |
|---|---|
| `help` | Lists these commands |
| `dump` | Replies with the raw JSON of the activity you just sent |
| `raw` | Sends a message built as a raw JSON object, bypassing the typed models |
| `custom` | Tries to send a custom activity type and reports the connector's answer verbatim |
| anything else | Echoes back, tagged with the channel it arrived on |

## What this sample teaches

### Raw sends must supply `from` yourself

`ctx.send()` goes through teams.ts's `ActivitySender`, which fills in `from` and
`conversation` for you. Posting raw JSON to the connector skips all of that, and the send
fails:

```
400 MissingProperty: The 'Activity.From' field is required
```

`postRaw` in `src/index.ts` adds `from` back from `ctx.activity.recipient`. That single field
is the practical cost of dropping to raw JSON.

### Custom activity types are accepted on some channels and rejected on others

`custom` sends `{"type": "vanilla/customActivity", ...}`. The result depends entirely on the
channel, not on the SDK:

| Channel | Result |
|---|---|
| Direct Line / Web Chat | **200** — accepted, and it appears in the transcript |
| Microsoft Teams | **400 BadArgument — "Unknown activity type"** |

Teams accepts only `message` and `typing` on `POST /v3/conversations/{id}/activities`. Other
outbound operations do exist on Teams, but as REST verbs rather than activity types —
`PUT /activities/{id}` to edit, `DELETE /activities/{id}` to delete, and
`PUT|DELETE /activities/{id}/reactions/{type}` to react.

### teams.ts is channel-agnostic

Nothing here is gated on `channelId`. `ctx.api` is rebuilt per activity from the inbound
`serviceUrl`, so sends go to the right place regardless of channel. Run `dump` on Direct Line
and again on email to see how much the payload shape changes while the handler stays identical.

## Testing

**Direct Line** — the quickest loop, and scriptable:

```bash
python3 tools/webchat/dl_test.py help dump raw custom "hello there"
```

**Email** — send a plain email to the address on the bot's email channel
(`az bot email show -n <botAppId> -g <rg>`), with the command in the body. Replies come back
as email. Adaptive cards flatten to static images on this channel, which is why this sample
sticks to plain text.

**Teams** — install the app and message it directly.

## How this differs from `teams-sdk-sample`

| | `teams-sdk-vanilla` | `teams-sdk-sample` |
|---|---|---|
| Stack | teams.ts only | Agents SDK + teams.ts via `TeamsSdkMiddleware` |
| Entry point | `new App()` + `app.start()` | `AgentApplication` + express + `useTeamsSdk` |
| Config | `CLIENT_ID` / `CLIENT_SECRET` / `TENANT_ID` env vars | `loadAuthConfigFromEnv` |
| Auth | none | two Graph connections, per-handler tokens |
| Focus | raw activities, channel behaviour | routing, fallthrough, multi-auth, multichannel |
