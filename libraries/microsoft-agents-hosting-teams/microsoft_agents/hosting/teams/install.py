"""One-call setup for Teams support.

Developer:
    from microsoft_agents.hosting.core import AgentApplication
    from microsoft_agents.hosting.teams import install_teams, TeamsHandlers

    app = AgentApplication()
    install_teams(app)
    teams = TeamsHandlers(app)
"""

from __future__ import annotations

from microsoft_agents.hosting.core.app.agent_application import AgentApplication

from .teams_middleware import TeamsMiddleware


def install_teams(app: AgentApplication) -> TeamsMiddleware:
    """Register Teams plumbing on the application's adapter.

    Returns the middleware instance so callers may inspect / extend it.
    """
    middleware = TeamsMiddleware()
    app.adapter.use(middleware)
    return middleware
