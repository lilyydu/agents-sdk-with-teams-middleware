# Vanilla Teams SDK sample

Plain **teams.py** — it exists to demonstrate two things:

1. **Raw JSON activities.** Sending and inspecting activities as plain dictionaries rather
   than typed models, and seeing exactly what the connector answers.
2. **Non-Teams channels.** teams.py never inspects `channelId`, so the same handlers run on
   Direct Line and email. `dump` makes the differences visible.

## Setup

```bash
cd python/test_samples/teams_sdk_vanilla
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
cp .env.example .env      # then fill in CLIENT_ID / CLIENT_SECRET / TENANT_ID
.venv/bin/python app.py
```

teams.py reads `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID` straight from the environment —
there is no configuration object and no `load_configuration_from_env`.

The app listens on **3979** so it can run alongside `teams_sdk_sample` on 3978. Point a tunnel
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
| `raw` | Sends a message built as a raw JSON dict, bypassing the typed models |
| `custom` | Tries to send a custom activity type and reports the connector's answer verbatim |
| anything else | Echoes back, tagged with the channel it arrived on |

## What this sample teaches

### Raw sends must supply `from` yourself

`ctx.send()` goes through teams.py's `ActivitySender`, which fills in `from` and
`conversation` for you. Posting raw JSON to the connector skips all of that, and the send
fails:

```
400 MissingProperty: The 'Activity.From' field is required
```

`_post_raw` in `app.py` adds `from` back from `ctx.activity.recipient`. That single field is
the practical cost of dropping to raw JSON.

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

Note the asymmetry: teams.py will happily *send* a custom type, but it cannot *receive* one.
Inbound activities are validated against a closed tagged union, and an unknown `type` raises
past the SDK's error handling and surfaces as **HTTP 500**. In practice this is unreachable on
Teams — Teams only ever emits its own known types — but it does mean the SDK would break on an
activity type Microsoft adds in the future.

### teams.py is channel-agnostic

Nothing here is gated on `channelId`. `ctx.api` is rebuilt per activity from the inbound
`serviceUrl`, so sends go to the right place regardless of channel. Run `dump` on Direct Line
and again on email to see how much the payload shape changes while the handler stays identical.

## Testing

**Direct Line** — the quickest loop, and scriptable:

```bash
python3 tools/webchat/dl_test.py help dump raw custom "hello there"
```

**Email** — send a plain email to the address on the bot's email channel
(`az bot email show -n <botAppId> -g <rg>`), with the command as the subject or body. Replies
come back as email. Adaptive cards flatten to static images on this channel, which is why this
sample sticks to plain text.

**Teams** — install the app and message it directly.

## How this differs from `teams_sdk_sample`

| | `teams_sdk_vanilla` | `teams_sdk_sample` |
|---|---|---|
| Stack | teams.py only | Agents SDK + teams.py via `TeamsSDKMiddleware` |
| Entry point | `App()` + `app.start()` | `AgentApplication` + aiohttp + `use_teams_sdk` |
| Config | `CLIENT_ID` / `CLIENT_SECRET` / `TENANT_ID` env vars | `load_configuration_from_env` |
| Auth | none | two Graph connections, per-handler tokens |
| Focus | raw activities, channel behaviour | routing, fallthrough, multi-auth, multichannel |
