"""Phase 3: + TeamsDeepExtension. Rich Teams handlers (additive).

The generic `@app.message()` still runs for ordinary Teams text messages
because the DeepExtension's can_handle returns False for activities with no
registered Teams route. Teams-specific things (invokes, reactions, ...) are
owned by the DeepExtension.
"""

from microsoft_agents.hosting.core import AgentApplication, TurnContext
from microsoft_agents.hosting.core.app.state import TurnState
from microsoft_agents.hosting.teams.teams_channel_adapter import TeamsChannelAdapter
from microsoft_agents.hosting.teams.teams_context import TeamsContext
from microsoft_agents.hosting.teams.teams_deep_extension import TeamsDeepExtension

app = AgentApplication[TurnState]()

teams = TeamsDeepExtension()
app.use(TeamsChannelAdapter(app_id="...", app_password="..."))
app.use(teams)


@app.message()
async def on_message(ctx: TurnContext, state: TurnState) -> None:
    # Still cross-platform — runs for ALL channels including Teams when no
    # Teams-specific route claims the turn.
    await ctx.send_activity(f"Echo: {ctx.activity.text}")


# ---- Layer-2 typed sugar ----

@teams.message_extension.on_query("search")
async def me_search(tctx: TeamsContext, query) -> dict:
    return {
        "composeExtension": {
            "type": "result",
            "attachmentLayout": "list",
            "attachments": [],  # build attachments from query.parameters
        }
    }


@teams.task_module.on_fetch()
async def open_dialog(tctx: TeamsContext, req) -> dict:
    return {"task": {"type": "continue", "value": {"title": "Hello"}}}


# ---- Layer-1 escape hatch — works the day Teams ships a new invoke type ----

@teams.on_invoke("adaptiveCard/action")
async def card_action(tctx: TeamsContext, payload: dict) -> dict:
    return {"statusCode": 200, "type": "application/vnd.microsoft.card.adaptive", "value": {}}
