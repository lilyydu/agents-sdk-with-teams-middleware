# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

import asyncio
import logging
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

from teams_sdk import (
    agent_sdk_turn_context,
    use_teams_sdk,
)

# teams.py — owns every Teams turn with a matching route
from microsoft_teams.apps import ActivityContext
from microsoft_teams.api import (
    AdaptiveCardActionMessageResponse,
    AdaptiveCardInvokeActivity,
    ConversationUpdateActivity,
    MeetingEndEventActivity,
    MeetingStartEventActivity,
    MessageActivity,
    MessageActivityInput,
    MessageReactionActivity,
    MessageSubmitActionInvokeActivity,
    TaskFetchInvokeActivity,
    TaskModuleContinueResponse,
    TaskModuleInvokeResponse,
    TaskModuleMessageResponse,
    TaskSubmitInvokeActivity,
)
from microsoft_teams.api.clients.api_client import ApiClient
from microsoft_teams.api.models.account import Account
from microsoft_teams.api.models.attachment import AdaptiveCardAttachment, card_attachment
from microsoft_teams.api.models.entity import CitationAppearance
from microsoft_teams.api.models.entity.citation_entity import CitationUsageInfo
from microsoft_teams.api.models.task_module import CardTaskModuleTaskInfo

from cards import help_card, ping_card, task_form_card, task_launcher_card


logging.basicConfig(level=logging.INFO)
log = logging.getLogger("rfc-sample")

# ──────────────────────────── Bootstrap ────────────────────────────

load_dotenv(path.join(path.dirname(__file__), ".env"))
agents_sdk_config = load_configuration_from_env(environ)

STORAGE = MemoryStorage()
CONNECTION_MANAGER = MsalConnectionManager(**agents_sdk_config)
ADAPTER = CloudAdapter(connection_manager=CONNECTION_MANAGER)

# ─── Agents SDK side ──────────────────────────────────────────────
AGENT_SDK_APP = AgentApplication[TurnState](
    options=ApplicationOptions(storage=STORAGE, adapter=ADAPTER),
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

@TEAMS_APP.on_message_pattern("help")
async def _help(ctx: ActivityContext[MessageActivity]):
    """List the commands this sample understands."""
    await ctx.send(MessageActivityInput().add_card(help_card()))


@TEAMS_APP.on_message_pattern("cards")
async def _cards(ctx: ActivityContext[MessageActivity]):
    """Send an Adaptive Card. Pressing the button fires an Action.Execute invoke
    routed to ``_on_card_action`` below; teams.py wraps the typed response in
    an ``InvokeResponse`` and ``TeamsSDKMiddleware`` propagates it back through
    the Agents SDK adapter so the HTTP layer writes the correct response."""
    await ctx.send(MessageActivityInput().add_card(ping_card()))


@TEAMS_APP.on_message_pattern("citation")
async def _citation(ctx: ActivityContext[MessageActivity]):
    """Send a message with a citation, sensitivity label, AI-generated label, and feedback affordance."""
    activity = (
        MessageActivityInput(text="Here is a response with citations [1].")
        .add_citation(
            1,
            CitationAppearance(
                name="Teams SDK Documentation",
                abstract="Documentation for the Microsoft Teams SDK.",
                url="https://microsoft.github.io/teams-sdk/welcome",
                usage_info=CitationUsageInfo(
                    at_id="sensitivity-1",
                    name="Confidential",
                    description="This information is confidential and for internal use only.",
                ),
            ),
        )
        .add_ai_generated()
        .add_feedback()
    )
    await ctx.send(activity)


@TEAMS_APP.on_message_pattern("stream")
async def _stream(ctx: ActivityContext[MessageActivity]):
    """Streaming response: informative status update → chunked text → finalize."""
    ctx.stream.update("Thinking…")
    ctx.stream.emit("Streaming is a powerful feature ")
    ctx.stream.emit("for sending long responses ")
    ctx.stream.emit("incrementally.")
    await ctx.stream.close()


@TEAMS_APP.on_message_pattern("react")
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


@TEAMS_APP.on_message_pattern("quote")
async def _quote(ctx: ActivityContext[MessageActivity]):
    """Reply to the user's message with a quoted reply (auto-quotes inbound)."""
    await ctx.reply("Quoting your message!")


@TEAMS_APP.on_message_pattern("targeted")
async def _targeted(ctx: ActivityContext[MessageActivity]):
    """Send a targeted (ephemeral) message visible only to the sender."""
    sender = ctx.activity.from_account
    targeted_msg = (
        MessageActivityInput(text="👁️ This message is only visible to you.")
        .with_recipient(Account(id=sender.id, name=sender.name), is_targeted=True)
    )
    await ctx.send(targeted_msg)


@TEAMS_APP.on_message_pattern("proactive")
async def _proactive(ctx: ActivityContext[MessageActivity]):
    """Fire-and-forget delayed proactive message."""
    conv_id = ctx.activity.conversation.id
    await ctx.send("Proactive message coming in ~3s…")

    async def _later():
        await asyncio.sleep(3)
        try:
            await TEAMS_APP.send(
                conv_id,
                MessageActivityInput().add_text("📣 Proactive message!"),
            )
        except Exception:
            log.exception("proactive: TEAMS_APP.send failed")

    asyncio.create_task(_later())


@TEAMS_APP.on_message_pattern("task")
async def _task(ctx: ActivityContext[MessageActivity]):
    """Send a card whose button opens a task module (task/fetch → task/submit)."""
    await ctx.send(MessageActivityInput().add_card(task_launcher_card()))


@TEAMS_APP.on_message_pattern("turn context")
async def _turn_context(ctx: ActivityContext[MessageActivity]):
    """Send via Teams SDK *and* Agents SDK from the same teams.py handler.

    ``agent_sdk_turn_context()`` returns the live Agents SDK ``TurnContext`` that
    ``TeamsSDKMiddleware`` built for this turn, so this handler can call into
    the Agents SDK outbound pipeline (and read/write ``turn_state``) without
    spinning up a second context."""
    agent_sdk_ctx = agent_sdk_turn_context()
    await ctx.send("[Teams SDK] Sending via teams.py ActivityContext…")
    await agent_sdk_ctx.send_activity(
        "[Agent SDK] Sending via Agents SDK TurnContext "
        "from inside a teams.py handler."
    )


# ─── Adaptive Card invoke handler ─────────────────────────────────

@TEAMS_APP.on_card_action_execute
async def _on_card_action(ctx: ActivityContext[AdaptiveCardInvokeActivity]):
    """Catch-all INVOKE handler for any Action.Execute on an Adaptive Card."""
    action = ctx.activity.value.action
    log.info("INVOKE adaptiveCard/action verb=%s data=%s", action.verb, action.data)
    await ctx.send(
        f"[Teams SDK] Adaptive Card action received. Data: {action.data}"
    )
    return AdaptiveCardActionMessageResponse(value="Action handled.")


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


@TEAMS_APP.on_message_submit_feedback
async def _on_feedback(ctx: ActivityContext[MessageSubmitActionInvokeActivity]):
    await ctx.send("[Teams SDK] Thanks for your feedback!")


@TEAMS_APP.on_conversation_update
async def _on_conversation_update(ctx: ActivityContext[ConversationUpdateActivity]):
    """Welcome new members."""
    added = ctx.activity.members_added or []
    bot_id = ctx.activity.recipient.id if ctx.activity.recipient else None
    if any(m.id != bot_id for m in added):
        await ctx.send(
            "Hello from Teams SDK! Type **help** to see available commands."
        )


@TEAMS_APP.on_meeting_start
async def _on_meeting_start(ctx: ActivityContext[MeetingStartEventActivity]):
    await ctx.send("[Teams SDK] Meeting has started!")


@TEAMS_APP.on_meeting_end
async def _on_meeting_end(ctx: ActivityContext[MeetingEndEventActivity]):
    await ctx.send("[Teams SDK] Meeting has ended!")


# ════════════════════ AGENT_SDK_APP — fallthrough + "agents sdk *" commands ════════════════════
# These fire for Teams activities that have no matching teams.py route (TeamsSDKMiddleware
# falls through) and for any non-Teams channel.

@AGENT_SDK_APP.conversation_update("membersAdded")
async def _agent_sdk_welcome(context: TurnContext, _state: TurnState):
    """Welcome from the Agents SDK side. Note: when TEAMS_APP also has an
    on_conversation_update handler that matches, the middleware short-circuits
    to teams.py and this handler never runs for Teams turns. It's still useful
    for non-Teams channels."""
    added = context.activity.members_added or []
    bot_id = context.activity.recipient.id if context.activity.recipient else None
    if any(m.id != bot_id for m in added):
        await context.send_activity(
            "[Agent SDK] Welcome! (This is the Agents SDK welcome path.)"
        )


@AGENT_SDK_APP.message("agents sdk react")
async def _agents_sdk_react(context: TurnContext, _state: TurnState):
    """Reach into teams.py's API client from an Agents SDK handler.

    ``TEAMS_APP.api`` is pinned to the service URL provided at App construction,
    so for handlers driven by the Agents SDK (whose activities may arrive on a
    different service URL) we build a per-turn ``ApiClient`` against the inbound
    ``context.activity.service_url`` while reusing the shared HTTP client."""
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


@AGENT_SDK_APP.message("agents sdk proactive")
async def _agents_sdk_proactive(context: TurnContext, _state: TurnState):
    """Send a proactive-style message via teams.py's API client.

    ``TEAMS_APP.api`` is pinned to the service URL provided at App construction,
    so for handlers driven by the Agents SDK (whose activities may arrive on a
    different service URL) we build a per-turn ``ApiClient`` against the inbound
    ``context.activity.service_url`` while reusing the shared HTTP client."""
    conv_id = context.activity.conversation.id
    api = ApiClient(service_url=context.activity.service_url, options=TEAMS_APP.api.http)
    await api.conversations.activities(conv_id).create(
        MessageActivityInput().add_text(
            "[Teams SDK] Proactive message triggered from an Agents SDK handler!"
        )
    )


@AGENT_SDK_APP.message("agents sdk citation")
async def _agents_sdk_citation(context: TurnContext, _state: TurnState):
    """Build a Teams citation activity in Agents SDK code, send via teams.py.

    Uses a per-turn ``ApiClient`` so the send targets the inbound service URL."""
    conv_id = context.activity.conversation.id
    teams_message = (
        MessageActivityInput(
            text="Agent SDK handler built this citation using Teams SDK types [1]."
        )
        .add_citation(
            1,
            CitationAppearance(
                name="Agent SDK Documentation",
                abstract="Documentation for the Microsoft 365 Agents SDK.",
                url="https://learn.microsoft.com/en-us/microsoft-365/agents-sdk/",
                usage_info=CitationUsageInfo(
                    at_id="sensitivity-1",
                    name="Confidential",
                    description="This information is confidential and for internal use only.",
                ),
            ),
        )
        .add_ai_generated()
        .add_feedback()
    )
    api = ApiClient(service_url=context.activity.service_url, options=TEAMS_APP.api.http)
    await api.conversations.activities(conv_id).create(teams_message)


@AGENT_SDK_APP.activity("message")
async def _echo(context: TurnContext, _state: TurnState):
    """Default echo fallthrough. Fires when no teams.py route matches and
    none of the ``agents sdk *`` commands above matched either."""
    text = (context.activity.text or "").strip()
    await context.send_activity(f"[Agent SDK] You said: {text}")


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
