"""Phase 1 — plumbing only.

Developer writes 100% generic AgentApplication code. They never import a
Teams-specific decorator. One call (`install_teams`) makes every send
Teams-correct: right connector client, notification banner, mention
formatting, card transforms.

Run on Slack / WebChat / Copilot Studio — exact same code, no Teams logic
runs (middleware no-ops when channel_id != 'msteams').
"""

from microsoft_agents.hosting.core import AgentApplication
from microsoft_agents.hosting.teams import install_teams

app = AgentApplication()
install_teams(app)


@app.message("hello")
async def hello(ctx, state):
    await ctx.send_activity("Hi there!")
