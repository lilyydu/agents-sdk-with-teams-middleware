"""Phase 2 — invokes via the escape hatch.

Use when Teams ships a new invoke type and the Teams package hasn't yet
added a typed decorator. Works day-1: ``on_invoke(name)`` matches on
``activity.name`` and you handle the raw value yourself.
"""

from microsoft_agents.hosting.core import AgentApplication
from microsoft_agents.hosting.teams import install_teams, TeamsHandlers

app = AgentApplication()
install_teams(app)
teams = TeamsHandlers(app)


@teams.on_invoke("composeExtension/somethingNew")
async def something_new(ctx, value):
    # value is the raw activity.value; you decide the response body
    return {"customResponse": "ok", "echo": value}
