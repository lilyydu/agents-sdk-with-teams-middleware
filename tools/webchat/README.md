# Web Chat / Direct Line harness

A tiny Direct Line client for exercising the samples on a **non-Teams** channel, so you can
see the half of the middleware contract that Teams cannot show: activities that pass straight
through to the Agents SDK app.

Nothing here is sample-specific; it drives whichever sample is currently bound to the bot
registration's endpoint.

## Setup

Fetch the Direct Line secret for the Azure bot:

```powershell
az bot directline show --name <botId> --resource-group <rg> --with-secrets -o json
```

Then either export it or drop it in `.env` next to these files:

```text
DIRECTLINE_SECRET=<secret>
```

## Browser UI

```powershell
py tools\webchat\serve.py
```

This serves `index.html` plus a `/api/token` endpoint that exchanges the Direct Line *secret*
for a short-lived *token*, so the secret stays on the local process and never reaches the
browser.

## Scripted

```powershell
py tools\webchat\dl_test.py help channel "agents sdk react"
```

This sends each argument as a message and prints the replies, including Adaptive Card contents.
