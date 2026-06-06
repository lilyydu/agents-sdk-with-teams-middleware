"""Phase 1: Agents SDK only. Works on every channel (LCM)."""

from microsoft_agents.hosting.core import AgentApplication, TurnContext
from microsoft_agents.hosting.core.app.state import TurnState

app = AgentApplication[TurnState]()


@app.message()
async def on_message(ctx: TurnContext, state: TurnState) -> None:
    await ctx.send_activity(f"Echo: {ctx.activity.text}")
