"""Phase 2: + TeamsChannelAdapter. Plumbing only — no handler changes.

What you gain:
  * Outbound activities go through TeamsConnectorClient (Teams team-owned).
  * on_activity_sending applies Teams transforms (mentions, channelData, ...).
  * channelData is parsed and stashed on turn_state for cross-platform handlers.
"""

from microsoft_agents.hosting.core import AgentApplication, TurnContext
from microsoft_agents.hosting.core.app.state import TurnState
from microsoft_agents.hosting.teams.teams_channel_adapter import TeamsChannelAdapter

app = AgentApplication[TurnState]()
app.use(TeamsChannelAdapter(app_id="...", app_password="..."))


@app.message()
async def on_message(ctx: TurnContext, state: TurnState) -> None:
    # Identical handler to Phase 1 — still cross-platform.
    # On Teams, plumbing is now Teams-correct.
    await ctx.send_activity(f"Echo: {ctx.activity.text}")
