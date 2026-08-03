# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

import asyncio
import logging
import re
from os import environ, path

from aiohttp import web
from dotenv import load_dotenv

from microsoft_agents.activity import load_configuration_from_env
from microsoft_agents.authentication.msal import MsalConnectionManager
from microsoft_agents.hosting.aiohttp import (
    CloudAdapter,
    jwt_authorization_middleware,
    start_agent_process,
)
from microsoft_agents.hosting.core import (
    AgentApplication,
    MemoryStorage,
    TurnContext,
    TurnState,
)
from microsoft_agents.hosting.core.app import ApplicationOptions

from teams_sdk import is_teams_channel, use_teams_sdk

# teams.py — owns every Teams turn with a matching route
from microsoft_teams.apps import ActivityContext
from microsoft_teams.api import (
    MessageActivity,
    MessageActivityInput,
    MessageReactionActivity,
    TaskFetchInvokeActivity,
    TaskModuleContinueResponse,
    TaskModuleInvokeResponse,
    TaskModuleMessageResponse,
    TaskSubmitInvokeActivity,
)
from microsoft_teams.api.clients.api_client import ApiClient
from microsoft_teams.api.models.account import Account
from microsoft_teams.api.models.attachment import AdaptiveCardAttachment, card_attachment
from microsoft_teams.api.models.task_module import CardTaskModuleTaskInfo

from cards import help_card, task_form_card, task_launcher_card


logging.basicConfig(level=logging.INFO)
log = logging.getLogger("rfc-sample")


def _command(name: str) -> re.Pattern[str]:
    mention = r"(?:<at\b[^>]*>.*?</at>|@\S+)"
    return re.compile(
        rf"\s*(?:{mention}\s*)*{re.escape(name)}[ \t]*(?:\r?\n[\s\S]*)?$",
        re.IGNORECASE | re.DOTALL,
    )

# ──────────────────────────── Bootstrap ────────────────────────────

load_dotenv(path.join(path.dirname(__file__), ".env"))
agents_sdk_config = load_configuration_from_env(environ)

STORAGE = MemoryStorage()
CONNECTION_MANAGER = MsalConnectionManager(**agents_sdk_config)
ADAPTER = CloudAdapter(connection_manager=CONNECTION_MANAGER)

# ─── Agents SDK side ──────────────────────────────────────────────
AGENT_SDK_APP = AgentApplication[TurnState](
    options=ApplicationOptions(storage=STORAGE, adapter=ADAPTER),
    connection_manager=CONNECTION_MANAGER,
)


@AGENT_SDK_APP.error
async def _on_error(context: TurnContext, error: Exception):
    log.exception("Unhandled error: %s", error)
    await context.send_activity(f"⚠️ {type(error).__name__}: {error}")


# ─── teams.py side ────────────────────────────────────────────────
# One call: extracts credentials from CONNECTION_MANAGER, wires teams.py's
# outbound token callback to it, constructs the App, and installs
# TeamsSDKMiddleware on AGENT_SDK_APP.adapter so Teams turns are short-circuited.
TEAMS_APP = use_teams_sdk(AGENT_SDK_APP, CONNECTION_MANAGER)


# ════════════════════ TEAMS_APP — Teams SDK feature showcase ════════════════════

@TEAMS_APP.on_message_pattern(_command("help"))
async def _help(ctx: ActivityContext[MessageActivity]):
    """List the commands this sample understands."""
    await ctx.send(MessageActivityInput().add_card(help_card()))


@TEAMS_APP.on_message_pattern(_command("react"))
async def _react(ctx: ActivityContext[MessageActivity]):
    """Bot adds, then removes, an emoji reaction on its own message."""
    response = await ctx.send("React to this message! I'll add 👍 and remove it.")
    conv_id = ctx.activity.conversation.id
    try:
        await ctx.api.reactions.add(conv_id, response.id, "like")
        await asyncio.sleep(2)
        await ctx.api.reactions.delete(conv_id, response.id, "like")
    except Exception:
        log.exception("react: reactions API call failed")


@TEAMS_APP.on_message_pattern(_command("quote"))
async def _quote(ctx: ActivityContext[MessageActivity]):
    """Reply to the user's message with a quoted reply (auto-quotes inbound)."""
    await ctx.reply("Quoting your message!")


@TEAMS_APP.on_message_pattern(_command("targeted"))
async def _targeted(ctx: ActivityContext[MessageActivity]):
    """Send a targeted (ephemeral) message visible only to the sender."""
    sender = ctx.activity.from_
    targeted_msg = (
        MessageActivityInput(text="👁️ This message is only visible to you.")
        .with_recipient(Account(id=sender.id, name=sender.name), is_targeted=True)
    )
    await ctx.send(targeted_msg)


@TEAMS_APP.on_message_pattern(_command("task"))
async def _task(ctx: ActivityContext[MessageActivity]):
    """Send a card whose button opens a task module (task/fetch → task/submit)."""
    await ctx.send(MessageActivityInput().add_card(task_launcher_card()))



# ─── Task module handlers ─────────────────────────────────────────

@TEAMS_APP.on_dialog_open
async def _on_task_fetch(
    ctx: ActivityContext[TaskFetchInvokeActivity],
) -> TaskModuleInvokeResponse:
    """task/fetch — return a task module continue with an adaptive card."""
    return TaskModuleInvokeResponse(
        task=TaskModuleContinueResponse(
            value=CardTaskModuleTaskInfo(
                title="Sample Task Module",
                card=card_attachment(AdaptiveCardAttachment(content=task_form_card())),
            )
        )
    )


@TEAMS_APP.on_dialog_submit
async def _on_task_submit(
    ctx: ActivityContext[TaskSubmitInvokeActivity],
) -> TaskModuleInvokeResponse:
    """task/submit — acknowledge and close."""
    data = ctx.activity.value.data
    await ctx.send(f"[Teams SDK] Task module submitted. Data: {data}")
    return TaskModuleInvokeResponse(task=TaskModuleMessageResponse(value="Done."))


# ─── Other Teams events ───────────────────────────────────────────

@TEAMS_APP.on_message_reaction
async def _on_message_reaction(ctx: ActivityContext[MessageReactionActivity]):
    added = ctx.activity.reactions_added or []
    removed = ctx.activity.reactions_removed or []
    summary = (
        f"added={[r.type for r in added]} removed={[r.type for r in removed]}"
    )
    await ctx.send(f"[Teams SDK] Reactions: {summary}")



# ════════════════════ AGENT_SDK_APP — fallthrough + "agents sdk *" commands ════════════════════
# These fire for Teams activities that have no matching teams.py route (TeamsSDKMiddleware
# falls through) and for any non-Teams channel.


@AGENT_SDK_APP.message(_command("help"))
async def _help_non_teams(context: TurnContext, _state: TurnState):
    await context.send_activity(
        "[Agent SDK] Commands: help, channel, agents sdk react, agents sdk proactive.\n"
        "Teams-only extras (react, quote, targeted, task) need the Teams SDK routes."
    )


@AGENT_SDK_APP.message(_command("channel"))
async def _channel(context: TurnContext, _state: TurnState):
    via = (
        "Teams turn with no matching teams.py route → fell through"
        if is_teams_channel(context.activity)
        else "non-Teams channel → passed straight through"
    )
    await context.send_activity(
        f"[Agent SDK] channelId={context.activity.channel_id} ({via})"
    )


@AGENT_SDK_APP.message(_command("agents sdk react"))
async def _agents_sdk_react(context: TurnContext, _state: TurnState):

    if not is_teams_channel(context.activity):
        await context.send_activity(
            f"[Agent SDK] 'agents sdk react' needs the Teams reactions API; "
            f"channelId={context.activity.channel_id} returns 404 for it."
        )
        return
    response = await context.send_activity(
        "[Agent SDK] Adding then removing 👍 via teams.py API client…"
    )
    conv_id = context.activity.conversation.id
    api = ApiClient(service_url=context.activity.service_url, options=TEAMS_APP.api.http)
    try:
        await api.reactions.add(conv_id, response.id, "like")
        await asyncio.sleep(2)
        await api.reactions.delete(conv_id, response.id, "like")
    except Exception:
        log.exception("agents sdk react: reactions API call failed")


@AGENT_SDK_APP.message(_command("agents sdk proactive"))
async def _agents_sdk_proactive(context: TurnContext, _state: TurnState):
    conv_id = context.activity.conversation.id
    api = ApiClient(service_url=context.activity.service_url, options=TEAMS_APP.api.http)
    bot = context.activity.recipient
    outgoing = MessageActivityInput().add_text(
        "[Teams SDK] Proactive message triggered from an Agents SDK handler!"
    )
    outgoing.from_ = Account(id=bot.id, name=bot.name)
    await api.conversations.activities(conv_id).create(outgoing)


@AGENT_SDK_APP.activity("message")
async def _echo(context: TurnContext, _state: TurnState):
    """Fallthrough: no teams.py route and no command above matched."""
    text = (context.activity.text or "").strip()
    # Email bodies carry a signature and/or quoted thread; echoing all of it
    # back grows on every round trip.
    first_line = next((ln.strip() for ln in text.splitlines() if ln.strip()), "")
    if first_line != text:
        text = f"{first_line} […]"
    await context.send_activity(
        f"[Agent SDK] ({context.activity.channel_id}) You said: {text}"
    )


# ─────────────────────────── HTTP wiring ───────────────────────────

async def _entry_point(req: web.Request) -> web.Response:
    return await start_agent_process(req, req.app["agent_sdk_app"], req.app["adapter"])


if __name__ == "__main__":
    APP = web.Application(middlewares=[jwt_authorization_middleware])
    APP.router.add_post("/api/messages", _entry_point)
    APP["agent_sdk_app"] = AGENT_SDK_APP
    APP["adapter"] = ADAPTER
    APP["agent_configuration"] = CONNECTION_MANAGER.get_default_connection_configuration()

    web.run_app(
        APP,
        host=environ.get("HOST", "localhost"),
        port=int(environ.get("PORT", "3978")),
    )
