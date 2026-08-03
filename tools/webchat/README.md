# Web Chat / Direct Line harness

A tiny Direct Line client for exercising the samples on a **non-Teams** channel, so you can
see the half of the middleware contract that Teams can't show: activities that pass straight
through to the Agents SDK app.

Nothing here is sample-specific — it drives whichever sample is currently bound to the bot
registration's endpoint.

## Setup

Direct Line is already enabled on the Azure Bot registration. Fetch the secret:

```bash
az bot directline show --name <botName> --resource-group <rg> --with-secrets -o json
```

Then either export it or drop it in `.env` next to these files (gitignored):

```
DIRECTLINE_SECRET=<secret>
```

## Browser UI

```bash
python tools/webchat/serve.py        # http://localhost:3000
```

Serves `index.html` plus a `/api/token` endpoint that exchanges the Direct Line *secret* for
a short-lived *token*, so the secret stays in this process and never reaches the browser.
Use this when you want to see Adaptive Cards render.

## Scripted

```bash
python tools/webchat/dl_test.py help channel "agents sdk react"
```

Sends each argument as a message and prints the replies, including Adaptive Card contents —
card-only replies are otherwise invisible when you filter on `activity.text`.
