"""Phase 4: Override `message` on Teams only.

The DeepExtension takes over `@app.message()` for Teams traffic — its
`@teams.on_message()` route is registered, so can_handle returns True for
Teams text messages and the cross-platform handler is bypassed on Teams.
Other channels (Slack, generic) still hit the cross-platform handler.
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
async def cross_platform_message(ctx: TurnContext, state: TurnState) -> None:
    # Runs for Slack, generic, etc. — NOT for Teams now.
    await ctx.send_activity(f"[generic] {ctx.activity.text}")


@teams.on_message()
async def teams_message(tctx: TeamsContext, _payload=None) -> None:
    # Teams-only handler with rich context (streaming, graph, mention, ...).
    await tctx.send(f"[teams] {tctx.activity.text}")
