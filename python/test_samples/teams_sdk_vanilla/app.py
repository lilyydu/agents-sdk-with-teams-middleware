# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

"""Vanilla Teams SDK sample.

Two things this sample is here to show:

1. Raw JSON activities. Every command below sends and inspects activities as raw
   dictionaries instead of typed models, so you can see exactly what goes on the
   wire and exactly what Azure Bot Service says back.
2. Non-Teams channels. teams.py never checks channelId, so the same handlers run
   on email. `dump` is the easiest way to see how different a non-Teams payload is.
"""

import json
import logging
import re
from os import environ, path

from dotenv import load_dotenv

from microsoft_teams.api import MessageActivity, MessageActivityInput
from microsoft_teams.apps import ActivityContext, App
from microsoft_teams.apps.events import ErrorEvent

load_dotenv(path.join(path.dirname(__file__), ".env"))

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("vanilla")

APP = App()

MENTION_TAG = re.compile(r"<at\b[^>]*>.*?</at>", re.IGNORECASE | re.DOTALL)

HELP = (
    "Vanilla teams.py sample — no Agents SDK, no middleware.\n\n"
    "• help — this message\n"
    "• dump — reply with the raw JSON of the activity you just sent\n"
    "• raw — send a message built as a raw JSON dict, bypassing the typed models\n"
    "• custom — try to send a custom activity type and report what ABS answers\n"
    "• anything else — echo, tagged with the channel it arrived on"
)


def _clean(text: str | None) -> str:
    """Strip Teams @-mention markup so commands match in channels and group chats."""
    return MENTION_TAG.sub("", text or "").strip()


def _block(payload: object) -> str:
    return "```\n" + json.dumps(payload, indent=2, default=str) + "\n```"


async def _post_raw(
    ctx: ActivityContext[MessageActivity], payload: dict
) -> tuple[dict, int, str]:
    """POST a raw activity dict straight to the connector, skipping teams.py's models.

    Sending raw means giving up what ActivitySender normally does for you. The one thing
    it always fills in is `from`, and channels reject the activity without it, so it is
    added here. Returns the payload as actually sent, plus the connector's answer.

    ctx.api is rebuilt per activity from the inbound serviceUrl, so this stays correct
    on email as well as Teams.
    """
    bot = ctx.activity.recipient
    payload.setdefault("from", {"id": bot.id, "name": bot.name})

    url = f"{ctx.api.service_url}/v3/conversations/{ctx.activity.conversation.id}/activities"
    try:
        response = await ctx.api.http.post(url, json=payload)
        return payload, response.status_code, response.text
    except Exception as error:  # surfaced to the user rather than swallowed
        response = getattr(error, "response", None)
        status = getattr(response, "status_code", 0)
        body = getattr(response, "text", str(error))
        return payload, status, body


@APP.on_message
async def _on_message(ctx: ActivityContext[MessageActivity]):
    command = _clean(ctx.activity.text).lower()

    if command == "help":
        await ctx.send(HELP)

    elif command == "dump":
        # The typed model round-tripped back to the JSON the channel actually sent.
        raw = ctx.activity.model_dump(by_alias=True, exclude_none=True, mode="json")
        await ctx.send(f"Raw inbound activity ({ctx.activity.channel_id}):\n{_block(raw)}")

    elif command == "raw":
        sent, status, body = await _post_raw(
            ctx, {"type": "message", "text": "Sent as a raw JSON dict — no typed model."}
        )
        await ctx.send(f"Sent {_block(sent)}\nConnector answered **{status}** {body or '(empty body)'}")

    elif command == "custom":
        sent, status, body = await _post_raw(
            ctx, {"type": "vanilla/customActivity", "text": "custom activity type"}
        )
        await ctx.send(
            f"Tried to send a custom activity type: {_block(sent)}\n"
            f"Connector answered **{status}**: {body or '(empty body)'}"
        )

    else:
        await ctx.send(
            MessageActivityInput(
                text=f"({ctx.activity.channel_id}) You said: {_clean(ctx.activity.text)}"
            )
        )


@APP.event("error")
async def _on_error(event: ErrorEvent):
    log.error("Unhandled error: %s", event.error, exc_info=event.error)


if __name__ == "__main__":
    import asyncio

    asyncio.run(APP.start(int(environ.get("PORT", "3979"))))
